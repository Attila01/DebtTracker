# debt_manager_gui.py
# Purpose: Provides a graphical user interface for the Debt Management System.
#          Allows users to view, add, edit, and delete records for debts, accounts,
#          payments, goals, assets, revenue, and categories.
# Deploy in: C:\DebtTracker
# Version: 1.6 (2025-07-19) - Fixed UnboundLocalError: 'pk_col_name_from_schema'
#                            by ensuring primary key variable is always defined.

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os
import logging

# Import database manager functions and configuration
import debt_manager_db_manager as db_manager
from config import TABLE_SCHEMAS, LOG_FILE, LOG_DIR

# Import the new CSV synchronization function
from debt_manager_csv_sync import sqlite_to_csv # Renamed from excel_sync to csv_sync

# Ensure log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

# Configure logging
if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s: %(message)s',
                        handlers=[
                            logging.FileHandler(LOG_FILE, mode='a'),
                            logging.StreamHandler()
                        ])

class DebtManagerApp(tk.Tk):
    """Main application class for the Debt Management System GUI."""

    def __init__(self, *args, **kwargs):
        try:
            super().__init__(*args, **kwargs)
            logging.info("DebtManagerApp: Initializing GUI application.")
            self.title("Debt Management System")
            self.geometry("1000x650") # Increased height for better layout
            self.style = ttk.Style(self)
            self.style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'

            self.create_widgets()
            self.protocol("WM_DELETE_WINDOW", self.on_closing) # Handle window close event

            # Initial data load for the Dashboard tab
            self.load_dashboard_data()
            logging.info("DebtManagerApp: GUI initialized successfully.")
        except Exception as e:
            logging.critical(f"DebtManagerApp: CRITICAL ERROR during __init__: {e}", exc_info=True)
            messagebox.showerror("GUI Startup Error", f"An unrecoverable error occurred during GUI startup: {e}\nCheck DebugLog.txt for details.")
            self.destroy() # Ensure the window is destroyed if init fails

    def create_widgets(self):
        """Creates the main GUI widgets, including tabs and buttons."""
        try:
            logging.info("DebtManagerApp: Creating main widgets.")
            self.notebook = ttk.Notebook(self)
            self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

            self.tabs = {} # Dictionary to hold tab frames and their Treeviews

            # Dashboard Tab
            self.create_dashboard_tab()

            # Create tabs for each table defined in TABLE_SCHEMAS
            # Order them specifically
            tab_order = ['Debts', 'Accounts', 'Payments', 'Goals', 'Assets', 'Revenue', 'Categories']
            for table_name in tab_order:
                schema = TABLE_SCHEMAS[table_name]
                if table_name == 'Categories':
                    self.create_categories_tab()
                else:
                    self.create_data_tab(table_name, schema['primary_key'], schema['gui_fields'])

            # Add tabs to notebook in desired order
            self.notebook.add(self.tabs['Dashboard']['frame'], text='Dashboard')
            for tab_name in tab_order:
                self.notebook.add(self.tabs[tab_name]['frame'], text=tab_name)

            # Add a Reports tab
            self.create_reports_tab()
            logging.info("DebtManagerApp: Main widgets created.")
        except Exception as e:
            logging.critical(f"DebtManagerApp: CRITICAL ERROR during create_widgets: {e}", exc_info=True)
            messagebox.showerror("GUI Error", f"An error occurred while creating GUI widgets: {e}\nCheck DebugLog.txt for details.")
            raise # Re-raise to be caught by __init__

    def create_dashboard_tab(self):
        """Creates the Dashboard tab with snowball data and summary."""
        try:
            logging.info("DashboardTab: Creating dashboard tab.")
            frame = ttk.Frame(self.notebook)
            self.tabs['Dashboard'] = {'frame': frame}

            # Treeview for Snowball Data
            # Note: The columns here should match the output of get_table_data for Debts
            columns = ['Creditor', 'Amount', 'MinimumPayment', 'SnowballPayment', 'Status']
            self.tabs['Dashboard']['tree'] = ttk.Treeview(frame, columns=columns, show='headings')
            for col in columns:
                self.tabs['Dashboard']['tree'].heading(col, text=col)
                self.tabs['Dashboard']['tree'].column(col, width=150, anchor='center')
            self.tabs['Dashboard']['tree'].pack(pady=10, padx=10, fill='both', expand=True)

            # Scrollbar for Treeview
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tabs['Dashboard']['tree'].yview)
            self.tabs['Dashboard']['tree'].configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side='right', fill='y')
            self.tabs['Dashboard']['tree'].pack(side='left', fill='both', expand=True)

            # Load Data Button
            load_btn = ttk.Button(frame, text="Refresh Dashboard Data", command=self.load_dashboard_data)
            load_btn.pack(pady=5)

            # Summary Labels (placeholders for now)
            self.total_debt_label = ttk.Label(frame, text="Total Debt: $0.00", font=('Arial', 12, 'bold'))
            self.total_debt_label.pack(pady=2)
            self.total_savings_label = ttk.Label(frame, text="Total Savings: $0.00", font=('Arial', 12, 'bold'))
            self.total_savings_label.pack(pady=2)
            self.net_worth_label = ttk.Label(frame, text="Net Worth: $0.00", font=('Arial', 12, 'bold'))
            self.net_worth_label.pack(pady=2)

            # Projections Section
            ttk.Label(frame, text="Financial Projections:", font=('Arial', 12, 'bold')).pack(pady=10)
            self.projection_tree = ttk.Treeview(frame, columns=['Year', 'DebtRemaining', 'Savings', 'NetWorth'], show='headings')
            for col in ['Year', 'DebtRemaining', 'Savings', 'NetWorth']:
                self.projection_tree.heading(col, text=col)
                self.projection_tree.column(col, width=150, anchor='center')
            self.projection_tree.pack(pady=5, padx=10, fill='both', expand=True)

            generate_proj_btn = ttk.Button(frame, text="Generate Projection", command=self.generate_projection)
            generate_proj_btn.pack(pady=5)

            logging.info("Dashboard tab created.")
        except Exception as e:
            logging.critical(f"DashboardTab: CRITICAL ERROR during creation: {e}", exc_info=True)
            messagebox.showerror("GUI Error", f"An error occurred while creating the Dashboard tab: {e}\nCheck DebugLog.txt for details.")
            raise

    def load_dashboard_data(self):
        """Loads snowball data and updates summary labels on the Dashboard tab."""
        try:
            logging.info("DashboardTab: Loading dashboard data.")
            for item in self.tabs['Dashboard']['tree'].get_children():
                self.tabs['Dashboard']['tree'].delete(item)

            debts_df = db_manager.get_table_data('Debts')

            # Filter and sort for snowball
            snowball_debts = debts_df[
                (debts_df['Status'] != 'Paid Off') &
                (debts_df['Status'] != 'Closed')
            ].sort_values(by='Amount', ascending=True)

            for index, row in snowball_debts.iterrows():
                self.tabs['Dashboard']['tree'].insert('', 'end', values=(
                    row['Creditor'],
                    f"${row['Amount']:.2f}",
                    f"${row['MinimumPayment']:.2f}",
                    f"${row['SnowballPayment']:.2f}",
                    row['Status']
                ))
            logging.info(f"Loaded snowball data ({len(snowball_debts)} debts) for Dashboard.")

            # Update summary labels
            total_debt = debts_df[
                (debts_df['Status'] != 'Paid Off') &
                (debts_df['Status'] != 'Closed')
            ]['Amount'].sum() if not debts_df.empty else 0.0

            accounts_df = db_manager.get_table_data('Accounts')
            total_savings = accounts_df[
                (accounts_df['AccountType'].isin(['Checking', 'Savings', 'Investment'])) &
                (accounts_df['Status'].isin(['Open', 'Current', 'Active']))
            ]['Balance'].sum() if not accounts_df.empty else 0.0

            net_worth = total_savings - total_debt

            self.total_debt_label.config(text=f"Total Debt: ${total_debt:.2f}")
            self.total_savings_label.config(text=f"Total Savings: ${total_savings:.2f}")
            self.net_worth_label.config(text=f"Net Worth: ${net_worth:.2f}")
            logging.info("Dashboard summary labels updated.")

            # Also load initial projection data
            self.generate_projection()
        except Exception as e:
            logging.error(f"DashboardTab: Error loading dashboard data: {e}", exc_info=True)
            messagebox.showerror("Dashboard Error", f"Failed to load dashboard data: {e}\nCheck DebugLog.txt for details.")


    def generate_projection(self):
        """Generates and displays financial projections."""
        try:
            logging.info("Projection: Generating financial projection.")
            for item in self.projection_tree.get_children():
                self.projection_tree.delete(item)

            projection_df = db_manager.generate_financial_projection()
            if not projection_df.empty:
                for index, row in projection_df.iterrows():
                    self.projection_tree.insert('', 'end', values=(
                        row['Year'],
                        f"${row['DebtRemaining']:.2f}",
                        f"${row['Savings']:.2f}",
                        f"${row['NetWorth']:.2f}"
                    ))
                logging.info(f"Generated and displayed {len(projection_df)} rows of financial projection.")
            else:
                logging.warning("Projection: No projection data generated.")
                # messagebox.showinfo("Projection", "No data available to generate financial projection.") # Avoid too many popups
        except Exception as e:
            logging.error(f"Projection: Error generating projection: {e}", exc_info=True)
            messagebox.showerror("Projection Error", f"An error occurred while generating the projection: {e}\nCheck DebugLog.txt for details.")


    def create_data_tab(self, table_name, primary_key, gui_fields):
        """Creates a generic data tab with Treeview and CRUD buttons."""
        try:
            logging.info(f"DataTab: Creating tab for {table_name}.")
            frame = ttk.Frame(self.notebook)
            self.tabs[table_name] = {'frame': frame, 'primary_key': primary_key, 'gui_fields': gui_fields}

            # Treeview
            # Use GUI field names for display, but ensure primary_key is the first hidden column
            display_columns = [field['name'] for field in gui_fields]
            tree_columns = [primary_key] + display_columns # Internal column first, then display columns

            tree = ttk.Treeview(frame, columns=tree_columns, show='headings')
            self.tabs[table_name]['tree'] = tree

            # Configure headings and columns
            tree.column(primary_key, width=0, stretch=tk.NO) # Hide the primary key column
            tree.heading(primary_key, text="") # No heading text for PK

            for col_name in display_columns:
                tree.heading(col_name, text=col_name)
                tree.column(col_name, width=120, anchor='center')

            tree.pack(pady=10, padx=10, fill='both', expand=True)

            # Scrollbar for Treeview
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side='right', fill='y')
            tree.pack(side='left', fill='both', expand=True)

            # Buttons
            button_frame = ttk.Frame(frame)
            button_frame.pack(pady=5)

            load_btn = ttk.Button(button_frame, text="Load Data", command=lambda: self.load_table_data(table_name))
            load_btn.pack(side='left', padx=5)

            add_btn = ttk.Button(button_frame, text=f"Add {table_name}", command=lambda: self.open_add_edit_form(table_name))
            add_btn.pack(side='left', padx=5)

            edit_btn = ttk.Button(button_frame, text=f"Edit Selected {table_name}", command=lambda: self.open_add_edit_form(table_name, edit_mode=True))
            edit_btn.pack(side='left', padx=5)

            delete_btn = ttk.Button(button_frame, text=f"Delete Selected {table_name}", command=lambda: self.delete_selected_record(table_name))
            delete_btn.pack(side='left', padx=5)

            # Special buttons for specific tabs
            if table_name == 'Goals':
                update_progress_btn = ttk.Button(button_frame, text="Update Goal Progress", command=self.update_goal_progress)
                update_progress_btn.pack(side='left', padx=5)
            elif table_name == 'Accounts':
                update_balance_btn = ttk.Button(button_frame, text="Update Account Balances", command=self.update_account_balances)
                update_balance_btn.pack(side='left', padx=5)
            elif table_name == 'Debts':
                update_debt_btn = ttk.Button(button_frame, text="Update Debt Amounts", command=self.update_debt_amounts_and_payments)
                update_debt_btn.pack(side='left', padx=5)

            logging.info(f"DataTab: Tab for {table_name} created.")
        except Exception as e:
            logging.critical(f"DataTab: CRITICAL ERROR during creation of {table_name} tab: {e}", exc_info=True)
            messagebox.showerror("GUI Error", f"An error occurred while creating the {table_name} tab: {e}\nCheck DebugLog.txt for details.")
            raise

    def create_categories_tab(self):
        """Creates the Categories tab with Treeview and CRUD buttons."""
        try:
            logging.info("CategoriesTab: Creating categories tab.")
            frame = ttk.Frame(self.notebook)
            self.tabs['Categories'] = {'frame': frame, 'primary_key': 'CategoryID', 'gui_fields': TABLE_SCHEMAS['Categories']['gui_fields']}

            # Treeview
            display_columns = ['CategoryName']
            tree_columns = ['CategoryID'] + display_columns
            tree = ttk.Treeview(frame, columns=tree_columns, show='headings')
            self.tabs['Categories']['tree'] = tree

            tree.heading('CategoryName', text='CategoryName')
            tree.column('CategoryName', width=200, anchor='center')
            tree.column('CategoryID', width=0, stretch=tk.NO) # Hide ID

            tree.pack(pady=10, padx=10, fill='both', expand=True)

            # Scrollbar for Treeview
            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side='right', fill='y')
            tree.pack(side='left', fill='both', expand=True)

            # Buttons
            button_frame = ttk.Frame(frame)
            button_frame.pack(pady=5)

            load_btn = ttk.Button(button_frame, text="Load Data", command=lambda: self.load_table_data('Categories'))
            load_btn.pack(side='left', padx=5)

            add_btn = ttk.Button(button_frame, text="Add Category", command=lambda: self.open_add_edit_form('Categories'))
            add_btn.pack(side='left', padx=5)

            edit_btn = ttk.Button(button_frame, text="Edit Selected Category", command=lambda: self.open_add_edit_form('Categories', edit_mode=True))
            edit_btn.pack(side='left', padx=5)

            delete_btn = ttk.Button(button_frame, text="Delete Selected Category", command=lambda: self.delete_selected_record('Categories'))
            delete_btn.pack(side='left', padx=5)
            logging.info("Categories tab created.")
        except Exception as e:
            logging.critical(f"CategoriesTab: CRITICAL ERROR during creation: {e}", exc_info=True)
            messagebox.showerror("GUI Error", f"An error occurred while creating the Categories tab: {e}\nCheck DebugLog.txt for details.")
            raise

    def create_reports_tab(self):
        """Creates the Reports tab with options for generating reports."""
        try:
            logging.info("ReportsTab: Creating reports tab.")
            frame = ttk.Frame(self.notebook)
            self.tabs['Reports'] = {'frame': frame}

            ttk.Label(frame, text="Select Report Type:").pack(pady=5)
            report_types = ['Debt Summary', 'Daily Expenses', 'Snowball Progress', 'Account Balances', 'Financial Projection']
            self.report_type_combo = ttk.Combobox(frame, values=report_types, state='readonly')
            self.report_type_combo.set(report_types[0])
            self.report_type_combo.pack(pady=2)

            ttk.Label(frame, text="Select Date (for Daily Expenses):").pack(pady=5)
            self.report_date_entry = ttk.Entry(frame)
            self.report_date_entry.insert(0, datetime.now().strftime('%Y-%m-%d')) # Default to today
            self.report_date_entry.pack(pady=2)
            ttk.Label(frame, text="(YYYY-MM-DD)").pack()

            generate_btn = ttk.Button(frame, text="Generate Report", command=self.generate_report)
            generate_btn.pack(pady=10)

            self.report_output_text = tk.Text(frame, height=15, width=80, wrap='word')
            self.report_output_text.pack(pady=10)

            logging.info("Reports tab created.")
        except Exception as e:
            logging.critical(f"ReportsTab: CRITICAL ERROR during creation: {e}", exc_info=True)
            messagebox.showerror("GUI Error", f"An error occurred while creating the Reports tab: {e}\nCheck DebugLog.txt for details.")
            raise


    def load_table_data(self, table_name):
        """Loads data from the specified table into its Treeview."""
        try:
            logging.info(f"LoadData: Loading data for {table_name}.")
            tree = self.tabs[table_name]['tree']
            for item in tree.get_children():
                tree.delete(item)

            df = db_manager.get_table_data(table_name)
            if df.empty:
                logging.info(f"LoadData: No data found for {table_name}.")
                return

            # Determine the columns to display in the Treeview
            # This should match the 'tree_columns' defined in create_data_tab
            pk_col = self.tabs[table_name]['primary_key']
            display_columns = [field['name'] for field in TABLE_SCHEMAS[table_name]['gui_fields']]
            all_tree_columns = [pk_col] + display_columns

            # Ensure DataFrame has all columns expected by the Treeview, in order
            # Handle cases where a joined column (like CategoryName) might not be present if no join occurred
            for col in all_tree_columns:
                if col not in df.columns:
                    # If it's a 'display_col' from a combo box, look for its corresponding 'value_col' in df
                    found_mapping_col = False
                    for field_info in TABLE_SCHEMAS[table_name]['gui_fields']:
                        if field_info['name'] == col and field_info['type'] == 'combo' and 'source_value_col' in field_info:
                            if field_info['source_value_col'] in df.columns:
                                df[col] = df[field_info['source_value_col']] # Use the raw ID
                                found_mapping_col = True
                                break
                    if not found_mapping_col:
                        df[col] = None # Add missing columns as None

            # Reorder DataFrame columns to match the Treeview's expected order
            df = df[all_tree_columns]

            for index, row in df.iterrows():
                # Convert values to strings for display, especially for dates/floats
                values = []
                for col_name in all_tree_columns:
                    val = row[col_name]
                    # Special handling for foreign key display in Treeview
                    # Check if the column is a foreign key with a display column from another table
                    if col_name != pk_col: # Don't format the hidden PK
                        field_info = next((f for f in TABLE_SCHEMAS[table_name]['gui_fields'] if f['name'] == col_name), None)
                        if field_info and field_info.get('type') == 'combo' and 'source_display_col' in field_info and 'source_value_col' in field_info:
                            # If it's a foreign key, and db_manager returned both ID and Name (e.g., DebtName (ID 1)), use the name
                            # Or if db_manager directly returned the name (e.g., CategoryName)
                            # The 'value' from the row could be the ID or the joined name.
                            # We want the display name for the Treeview.
                            display_col_from_db = field_info['source_display_col']
                            value_col_from_db = field_info['source_value_col']

                            # Check if the row contains the explicit display column from the join
                            if display_col_from_db in row and row[display_col_from_db] is not None:
                                display_val = str(row[display_col_from_db])
                                if value_col_from_db != display_col_from_db and value_col_from_db in row and row[value_col_from_db] is not None:
                                     # Append ID in parentheses if different for clarity in Treeview
                                     display_val = f"{display_val} ({row[value_col_from_db]})"
                                values.append(display_val)
                                continue # Move to next column
                            # Fallback if display column not found or is None, use the raw value (ID)
                            elif val is not None:
                                values.append(str(val))
                                continue
                    values.append(str(val) if val is not None else '')

                tree.insert('', 'end', values=values)
            logging.info(f"LoadData: Loaded data for {table_name} ({len(df)} rows).")
        except Exception as e:
            logging.error(f"LoadData: Error loading data for {table_name}: {e}", exc_info=True)
            messagebox.showerror("Load Data Error", f"Failed to load data for {table_name}: {e}\nCheck DebugLog.txt for details.")

    def open_add_edit_form(self, table_name, edit_mode=False):
        """Opens a form for adding new records or editing existing ones."""
        try:
            logging.info(f"Form: Opening {'edit' if edit_mode else 'add'} form for {table_name}.")
            form_window = tk.Toplevel(self)
            form_window.title(f"{'Edit' if edit_mode else 'Add'} {table_name} Record")
            form_window.geometry("400x600")
            form_window.transient(self) # Make dialog modal

            entries = {}     # To store Tkinter widgets for each field

            selected_item_id = None
            current_record_data = {}
            # Ensure pk_col_name_from_schema is always defined
            pk_col_name_from_schema = self.tabs[table_name]['primary_key']

            if edit_mode:
                selected_items = self.tabs[table_name]['tree'].selection()
                if not selected_items:
                    messagebox.showwarning("Edit Record", "Please select a record to edit.")
                    form_window.destroy()
                    return
                selected_item_id = selected_items[0]
                # Get values from the Treeview item
                values = self.tabs[table_name]['tree'].item(selected_item_id, 'values')

                # Map values back to column names including the primary key
                display_cols = [field['name'] for field in TABLE_SCHEMAS[table_name]['gui_fields']]
                all_cols_ordered_in_treeview = [pk_col_name_from_schema] + display_cols

                # Create a dictionary from the values, using the order they appear in the treeview
                for i, col_name in enumerate(all_cols_ordered_in_treeview):
                    current_record_data[col_name] = values[i]

                logging.debug(f"Form: Editing record: {current_record_data}")


            row_num = 0
            for field_info in TABLE_SCHEMAS[table_name]['gui_fields']:
                field_name = field_info['name'] # This is the GUI display name
                field_type = field_info['type']

                ttk.Label(form_window, text=f"{field_name}:").grid(row=row_num, column=0, padx=5, pady=5, sticky='w')

                if field_type == 'text':
                    entry = ttk.Entry(form_window)
                    entries[field_name] = entry
                elif field_type == 'decimal':
                    entry = ttk.Entry(form_window)
                    entries[field_name] = entry
                elif field_type == 'date':
                    entry = ttk.Entry(form_window) # Simple entry for YYYY-MM-DD
                    entries[field_name] = entry
                elif field_type == 'combo':
                    combo = ttk.Combobox(form_window, state='readonly')
                    entries[field_name] = combo

                    combo_options = []
                    # Keep track of value to display map to properly pre-fill combo
                    value_to_display_map = {}

                    if field_info.get('allow_none', False):
                        combo_options.append("None") # Option for no selection
                        value_to_display_map[None] = "None"

                    if 'options' in field_info: # For hardcoded options (e.g., Status)
                        combo_options.extend(field_info['options'])
                        for opt in field_info['options']:
                            value_to_display_map[opt] = opt # For simple options, value and display are the same
                    elif 'source_table' in field_info: # For options from other tables (foreign keys)
                        source_tables = field_info['source_table']
                        source_display_cols = field_info['source_display_col']
                        source_value_cols = field_info['source_value_col']

                        # Ensure source_tables, source_display_cols, source_value_cols are lists for consistent iteration
                        if not isinstance(source_tables, list):
                            source_tables = [source_tables]
                            source_display_cols = [source_display_cols]
                            source_value_cols = [source_value_cols]

                        for i, src_tbl in enumerate(source_tables):
                            try:
                                df_source = db_manager.get_table_data(src_tbl)
                                if not df_source.empty:
                                    display_col = source_display_cols[i]
                                    value_col = source_value_cols[i]

                                    # Ensure display_col and value_col exist in the source DataFrame
                                    if display_col not in df_source.columns:
                                        logging.warning(f"Form: Source display column '{display_col}' not found in {src_tbl}. Skipping options for {field_name}.")
                                        continue
                                    if value_col not in df_source.columns:
                                        logging.warning(f"Form: Source value column '{value_col}' not found in {src_tbl}. Skipping options for {field_name}.")
                                        continue

                                    for idx, row in df_source.iterrows():
                                        display_val = row[display_col]
                                        value_val = row[value_col]

                                        # Format for display: "Name (ID)" or just "Name" if ID is the same as Name
                                        if value_col == display_col: # e.g. CategoryName in Categories table
                                            option_text = str(display_val)
                                            combo_options.append(option_text)
                                            value_to_display_map[value_val] = option_text
                                        else: # e.g. Creditor (DebtID)
                                            option_text = f"{display_val} ({value_val})"
                                            combo_options.append(option_text)
                                            value_to_display_map[value_val] = option_text

                            except Exception as e:
                                logging.error(f"Form: Error loading combo options from {src_tbl} for field {field_name}: {e}", exc_info=True)
                                # Continue to next source table or field if one fails
                    combo['values'] = combo_options
                    entries[field_name].value_to_display_map = value_to_display_map # Store map for pre-filling
                    combo.set(combo_options[0] if combo_options else "") # Set default if available

                entries[field_name].grid(row=row_num, column=1, padx=5, pady=5, sticky='ew')
                row_num += 1

            # Populate fields if in edit mode
            if edit_mode:
                for field_info in TABLE_SCHEMAS[table_name]['gui_fields']:
                    field_name = field_info['name'] # This is the GUI display name
                    field_type = field_info['type']

                    # The actual DB column name for the value (ID or raw text)
                    db_value_col_name = field_info.get('source_value_col', field_name)
                    # The name of the column that might contain the joined display text
                    db_display_col_name = field_info.get('source_display_col', field_name)

                    # Prioritize finding the explicit display column from get_table_data first
                    value_to_prefill = None
                    if db_display_col_name in current_record_data:
                        value_to_prefill = current_record_data[db_display_col_name]
                    # If not found, fall back to the actual value column from the DB
                    elif db_value_col_name in current_record_data:
                        value_to_prefill = current_record_data[db_value_col_name]

                    if value_to_prefill is not None:
                        if entries[field_name].winfo_class() == 'TCombobox':
                            combo_widget = entries[field_name]
                            # Try to match the actual DB value to its display string in the combo
                            if value_to_prefill in combo_widget.value_to_display_map:
                                combo_widget.set(combo_widget.value_to_display_map[value_to_prefill])
                            elif field_info.get('allow_none') and (str(value_to_prefill).lower() == 'none' or value_to_prefill == '' or value_to_prefill == 0):
                                combo_widget.set("None")
                            else:
                                # Fallback: if value not found in map, try to set raw value
                                combo_widget.set(str(value_to_prefill))
                                logging.warning(f"Form: Could not find display mapping for {field_name} (prefill value: {value_to_prefill}). Setting raw value.")

                        elif field_type == 'date':
                            # Format date from DB (which might include time) to YYYY-MM-DD for entry
                            try:
                                # Handles 'YYYY-MM-DD HH:MM:SS.f' or 'YYYY-MM-DD'
                                dt_obj = datetime.strptime(str(value_to_prefill).split(' ')[0], '%Y-%m-%d')
                                entries[field_name].delete(0, tk.END)
                                entries[field_name].insert(0, dt_obj.strftime('%Y-%m-%d'))
                            except ValueError:
                                entries[field_name].delete(0, tk.END)
                                entries[field_name].insert(0, str(value_to_prefill)) # Insert as is if format fails
                                logging.warning(f"Form: Invalid date format for pre-fill: {value_to_prefill}. Inserting as raw string.")
                        else:
                            entries[field_name].delete(0, tk.END)
                            entries[field_name].insert(0, str(value_to_prefill))
                    else:
                        logging.warning(f"Form: No data found for field '{field_name}' (DB cols: '{db_display_col_name}', '{db_value_col_name}') in current_record_data for {table_name}. Cannot pre-fill.")


            # Fix for NameError: pk_col in lambda
            # Capture pk_col_name_from_schema in the lambda's default arguments
            # This ensures pk_col_name_from_schema is evaluated at definition time, not execution time.
            save_button = ttk.Button(
                form_window,
                text="Save",
                command=lambda pk=pk_col_name_from_schema: self.save_form_data(
                    form_window,
                    table_name,
                    entries,
                    edit_mode,
                    current_record_data.get(pk) if current_record_data else None
                )
            )
            save_button.grid(row=row_num, column=0, columnspan=2, pady=10)

            form_window.grab_set() # Make it modal
            self.wait_window(form_window) # Wait for the form to close
            logging.info(f"Form: {'Edit' if edit_mode else 'Add'} form for {table_name} opened and closed.")
        except Exception as e:
            logging.error(f"Form: Error opening add/edit form for {table_name}: {e}", exc_info=True)
            messagebox.showerror("Form Error", f"Failed to open add/edit form for {table_name}: {e}\nCheck DebugLog.txt for details.")
            if 'form_window' in locals() and form_window.winfo_exists():
                form_window.destroy()

    def save_form_data(self, form_window, table_name, entries, edit_mode, record_id=None):
        """Saves data from the add/edit form to the database."""
        try:
            logging.info(f"SaveForm: Saving data for {table_name} (edit_mode={edit_mode}).")
            data_to_save = {}
            for field_name, entry_widget in entries.items():
                value = entry_widget.get().strip()
                field_info = next(f for f in TABLE_SCHEMAS[table_name]['gui_fields'] if f['name'] == field_name)
                field_type = field_info['type']

                # Determine the actual DB column name to save to
                db_col_name = field_name # Default to GUI field name, used for text/decimal/date

                # If it's a combo, we need to save the source_value_col (ID) to the DB.
                # The GUI field_name here is just the display name (e.g., 'Category', 'Account').
                if field_type == 'combo' and 'source_value_col' in field_info:
                    db_col_name = field_info['source_value_col'] # Use the actual DB column name for the ID

                    # Handle "None" selection or empty string for nullable foreign keys
                    if value.lower() == "none" or value == "":
                        value = None
                    else:
                        # For comboboxes, extract the value to be saved to the database.
                        # This could be the ID in parentheses, or the raw string if value_col == display_col.
                        if '(' in value and ')' in value and field_info['source_value_col'] != field_info['source_display_col']:
                            # Extract ID from "Name (ID)" format
                            try:
                                value = int(value.split('(')[-1].replace(')', ''))
                            except ValueError:
                                messagebox.showerror("Input Error", f"Invalid ID format in combo box for {field_name}. Please select a valid item or 'None'.")
                                logging.warning(f"SaveForm: Invalid ID format for {field_name}: {value}")
                                return
                        # If value_col == display_col (e.g., CategoryName for Assets.Category), keep as string
                        # Or if it's a simple combo (options, not source_table), keep as string
                        # In these cases, 'value' is already the correct string to save
                        pass

                # Type conversion based on the column's type in TABLE_SCHEMAS for the *actual DB column*
                db_schema_col_info = next(c for c in TABLE_SCHEMAS[table_name]['columns'] if c['name'] == db_col_name)
                db_column_type = db_schema_col_info['type']

                if value is None:
                    data_to_save[db_col_name] = None
                elif db_column_type == 'REAL':
                    try:
                        data_to_save[db_col_name] = float(value)
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid numeric value for {field_name}.")
                        logging.warning(f"SaveForm: Invalid numeric value for {field_name}: {value}")
                        return
                elif db_column_type == 'INTEGER':
                    try:
                        data_to_save[db_col_name] = int(value)
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid integer value for {field_name}.")
                        logging.warning(f"SaveForm: Invalid integer value for {field_name}: {value}")
                        return
                elif db_column_type == 'TEXT':
                    # Special validation for date fields
                    if field_type == 'date':
                        try:
                            datetime.strptime(value, '%Y-%m-%d') # Validate format
                            data_to_save[db_col_name] = value # Store as string YYYY-MM-DD
                        except ValueError:
                            messagebox.showerror("Input Error", f"Invalid date format for {field_name}. Use YYYY-MM-DD.")
                            logging.warning(f"SaveForm: Invalid date format for {field_name}: {value}")
                            return
                    else:
                        data_to_save[db_col_name] = value
                else: # Fallback for unknown types, just save as is
                    data_to_save[db_col_name] = value

            success = False
            if edit_mode:
                success = db_manager.update_record(table_name, record_id, data_to_save)
                action = "updated"
            else:
                # For new records, handle 'OriginalAmount' and 'AmountPaid' for Debts
                # These might not be present in GUI fields but are needed for DB insertion
                if table_name == 'Debts':
                    if 'OriginalAmount' not in data_to_save or data_to_save['OriginalAmount'] is None:
                        data_to_save['OriginalAmount'] = data_to_save.get('Amount', 0.0)
                    if 'AmountPaid' not in data_to_save or data_to_save['AmountPaid'] is None:
                        data_to_save['AmountPaid'] = 0.0 # Initialize AmountPaid for new debts

                # For new Accounts, set Balance as InitialBalance if not provided
                if table_name == 'Accounts':
                    if 'Balance' not in data_to_save or data_to_save['Balance'] is None:
                        data_to_save['Balance'] = data_to_save.get('InitialBalance', 0.0)
                    if 'InitialBalance' not in data_to_save or data_to_save['InitialBalance'] is None:
                         data_to_save['InitialBalance'] = data_to_save.get('Balance', 0.0)

                success = db_manager.insert_record(table_name, data_to_save)
                action = "added"

            if success:
                messagebox.showinfo("Success", f"Record {action} successfully.")
                form_window.destroy()
                self.load_table_data(table_name) # Refresh the table data
                # Update related data if necessary
                if table_name in ['Payments', 'Revenue', 'Debts', 'Accounts', 'Goals']:
                    # Re-run all updates to ensure consistency across the board
                    # This might be overkill but ensures data integrity after any change
                    db_manager.update_debt_amounts_and_payments()
                    db_manager.update_account_balances()
                    db_manager.update_goal_progress()
                    self.load_dashboard_data() # Refresh dashboard for updated totals
                logging.info(f"SaveForm: Record {action} successfully for {table_name}.")
            else:
                messagebox.showerror("Error", f"Failed to {action} record. Check logs for details.")
                logging.error(f"SaveForm: Failed to {action} record for {table_name}.")
        except Exception as e:
            logging.error(f"SaveForm: Error saving form data for {table_name}: {e}", exc_info=True)
            messagebox.showerror("Save Error", f"An error occurred while saving data: {e}\nCheck DebugLog.txt for details.")


    def delete_selected_record(self, table_name):
        """Deletes the selected record from the specified table."""
        try:
            logging.info(f"DeleteRecord: Attempting to delete record(s) from {table_name}.")
            tree = self.tabs[table_name]['tree']
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showwarning("Delete Record", "Please select a record to delete.")
                return

            if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected record(s)?"):
                pk_col = self.tabs[table_name]['primary_key']
                success_count = 0
                for item_id in selected_items:
                    values = tree.item(item_id, 'values')
                    # The primary key is the first value in the `values` tuple because we put it first in `columns_with_pk`
                    record_id = values[0]

                    if db_manager.delete_record(table_name, pk_col, record_id):
                        success_count += 1
                    else:
                        messagebox.showerror("Delete Error", f"Failed to delete record with ID {record_id}.")

                if success_count > 0:
                    messagebox.showinfo("Success", f"{success_count} record(s) deleted successfully.")
                    self.load_table_data(table_name) # Refresh the table data
                    # Update related data if necessary
                    if table_name in ['Payments', 'Revenue', 'Debts', 'Accounts', 'Goals']:
                        db_manager.update_debt_amounts_and_payments()
                        db_manager.update_account_balances()
                        db_manager.update_goal_progress()
                        self.load_dashboard_data()
                    logging.info(f"DeleteRecord: {success_count} record(s) deleted successfully from {table_name}.")
                else:
                    messagebox.showwarning("Delete", "No records were deleted.")
        except Exception as e:
            logging.error(f"DeleteRecord: Error deleting record(s) from {table_name}: {e}", exc_info=True)
            messagebox.showerror("Delete Error", f"An error occurred while deleting records: {e}\nCheck DebugLog.txt for details.")

    def update_account_balances(self):
        """Calls the db_manager to update all account balances."""
        try:
            logging.info("UpdateBalances: Updating account balances.")
            db_manager.update_account_balances()
            messagebox.showinfo("Success", "Account balances updated successfully.")
            self.load_table_data('Accounts') # Refresh accounts tab
            self.load_dashboard_data() # Refresh dashboard totals
            logging.info("UpdateBalances: Account balances updated and GUI refreshed.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update account balances: {e}")
            logging.error(f"UpdateBalances: Error updating account balances from GUI: {e}", exc_info=True)

    def update_goal_progress(self):
        """Calls the db_manager to update all goal progress."""
        try:
            logging.info("UpdateGoals: Updating goal progress.")
            db_manager.update_goal_progress()
            messagebox.showinfo("Success", "Goal progress updated successfully.")
            self.load_table_data('Goals') # Refresh goals tab
            self.load_dashboard_data() # Refresh dashboard totals
            logging.info("UpdateGoals: Goal progress updated and GUI refreshed.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update goal progress: {e}")
            logging.error(f"UpdateGoals: Error updating goal progress from GUI: {e}", exc_info=True)

    def update_debt_amounts_and_payments(self):
        """Calls the db_manager to update debt amounts and payments."""
        try:
            logging.info("UpdateDebts: Updating debt amounts and payments.")
            db_manager.update_debt_amounts_and_payments()
            messagebox.showinfo("Success", "Debt amounts and payments updated successfully.")
            self.load_table_data('Debts') # Refresh debts tab
            self.load_dashboard_data() # Refresh dashboard totals
            logging.info("UpdateDebts: Debt amounts and payments updated and GUI refreshed.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update debt amounts and payments: {e}")
            logging.error(f"UpdateDebts: Error updating debt amounts and payments from GUI: {e}", exc_info=True)


    def generate_report(self):
        """Generates a report based on user selection and displays it."""
        try:
            logging.info("Reports: Generating report.")
            report_type = self.report_type_combo.get()
            report_date_str = self.report_date_entry.get()
            self.report_output_text.delete(1.0, tk.END) # Clear previous output

            df_report = pd.DataFrame()

            if report_type == 'Debt Summary':
                df_report = db_manager.get_table_data('Debts')
                # Exclude internal IDs from display if they are not part of gui_fields
                display_cols = [field['name'] for field in TABLE_SCHEMAS['Debts']['gui_fields']]
                # Ensure the display columns exist in the DataFrame before filtering
                df_report = df_report[[col for col in display_cols if col in df_report.columns]]
            elif report_type == 'Daily Expenses':
                try:
                    report_date = datetime.strptime(report_date_str, '%Y-%m-%d').strftime('%Y-%m-%d')
                except ValueError:
                    messagebox.showerror("Input Error", "Invalid date format. Please use YYYY-MM-DD.")
                    logging.warning(f"Reports: Invalid date format for Daily Expenses: {report_date_str}")
                    return
                payments_df = db_manager.get_table_data('Payments')
                df_report = payments_df[payments_df['PaymentDate'] == report_date]
                # Exclude internal IDs from display if they are not part of gui_fields
                display_cols = [field['name'] for field in TABLE_SCHEMAS['Payments']['gui_fields']]
                df_report = df_report[[col for col in display_cols if col in df_report.columns]]
            elif report_type == 'Snowball Progress':
                debts_df = db_manager.get_table_data('Debts')
                df_report = debts_df[
                    (debts_df['Status'] != 'Paid Off') &
                    (debts_df['Status'] != 'Closed')
                ].sort_values(by='Amount', ascending=True)
                # Ensure these columns exist
                df_report = df_report[['Creditor', 'Amount', 'MinimumPayment', 'SnowballPayment', 'Status']]
            elif report_type == 'Account Balances':
                df_report = db_manager.get_table_data('Accounts')
                # Exclude internal IDs from display if they are not part of gui_fields
                display_cols = [field['name'] for field in TABLE_SCHEMAS['Accounts']['gui_fields']]
                df_report = df_report[[col for col in display_cols if col in df_report.columns]]
            elif report_type == 'Financial Projection':
                df_report = db_manager.generate_financial_projection()

            if not df_report.empty:
                report_output = f"--- {report_type} Report ---\n\n"
                report_output += df_report.to_string(index=False)
                self.report_output_text.insert(tk.END, report_output)

                # Optionally save to CSV
                report_filename = f"{report_type.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                report_path = os.path.join(LOG_DIR, report_filename) # Using LOG_DIR for reports for now
                df_report.to_csv(report_path, index=False)
                messagebox.showinfo("Report Generated", f"Report saved to {report_path}")
                logging.info(f"Reports: {report_type} report generated and saved to {report_path}.")
            else:
                self.report_output_text.insert(tk.END, f"No data found for {report_type} report.")
                messagebox.showinfo("Report", f"No data found for {report_type} report.")
                logging.warning(f"Reports: No data found for {report_type} report.")

        except Exception as e:
            messagebox.showerror("Report Error", f"An error occurred while generating the report: {e}")
            logging.error(f"Reports: Error generating report: {e}", exc_info=True)

    def on_closing(self):
        """Handles actions to perform when the application window is closed."""
        logging.info("AppClosing: Handling application closing event.")
        if messagebox.askokcancel("Quit", "Do you want to quit the application?"):
            # Perform final sync to CSV before closing
            try:
                logging.info("Application is closing. Performing final sync to CSV...")
                db_manager.update_debt_amounts_and_payments() # Ensure latest calculated values are saved
                db_manager.update_account_balances()
                db_manager.update_goal_progress()
                # Call the new sqlite_to_csv function
                sqlite_to_csv()
                logging.info("Final sync completed. Destroying application window.")
            except Exception as e:
                logging.error(f"Error during final sync before closing: {e}", exc_info=True)
                messagebox.showwarning("Sync Error", f"An error occurred during final data sync to CSV: {e}\nData might not be fully saved. Check DebugLog.txt.")
            finally:
                self.destroy()
        else:
            logging.info("AppClosing: User cancelled quit.")

if __name__ == "__main__":
    # Ensure database is initialized before launching the GUI
    try:
        logging.info("Main: Attempting to initialize database before GUI launch.")
        # Call the specific initializer from db_init for comprehensive setup
        from debt_manager_db_init import initialize_database as db_initializer
        db_initializer()
        logging.info("Main: Database initialized for GUI launch.")
        app = DebtManagerApp()
        app.mainloop()
    except Exception as e:
        logging.critical(f"Main: CRITICAL ERROR during GUI startup from __main__: {e}", exc_info=True)
        messagebox.showerror("Startup Error", f"Failed to initialize database or launch GUI: {e}\nCheck DebugLog.txt for more details.")
