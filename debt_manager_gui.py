# debt_manager_gui.py
# Purpose: Provides a graphical user interface for the Debt Management System.
#          Allows users to view, add, edit, and delete records for debts, accounts,
#          payments, goals, assets, revenue, and categories.
# Deploy in: C:\DebtTracker
# Version: 1.1 (2025-07-19) - Added comprehensive error handling and logging
#          within the GUI application to catch early startup crashes.
#          Improved robustness for combo box data loading, especially for
#          multi-source fields like 'AllocatedTo'.

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os
import logging

# Import database manager functions and configuration
import debt_manager_db_manager as db_manager
from config import TABLE_SCHEMAS, LOG_FILE, LOG_DIR

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
            for table_name, schema in TABLE_SCHEMAS.items():
                if table_name == 'Categories': # Categories tab will be handled slightly differently
                    self.create_categories_tab()
                else:
                    self.create_data_tab(table_name, schema['primary_key'], schema['gui_fields'])

            # Re-order tabs if necessary (e.g., Dashboard first)
            self.notebook.add(self.tabs['Dashboard']['frame'], text='Dashboard')
            for tab_name in ['Debts', 'Accounts', 'Payments', 'Goals', 'Assets', 'Revenue', 'Categories']:
                if tab_name in self.tabs and tab_name != 'Dashboard':
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
            load_btn = ttk.Button(frame, text="Load Snowball Data", command=self.load_dashboard_data)
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
            logging.info(f"Loaded snowball data ({len(snowball_debts)} rows) for Dashboard.")

            # Update summary labels
            total_debt = debts_df[
                (debts_df['Status'] != 'Paid Off') &
                (debts_df['Status'] != 'Closed')
            ]['Amount'].sum()

            accounts_df = db_manager.get_table_data('Accounts')
            total_savings = accounts_df[
                (accounts_df['AccountType'].isin(['Checking', 'Savings', 'Investment'])) &
                (accounts_df['Status'].isin(['Open', 'Current', 'Active']))
            ]['Balance'].sum()

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
                messagebox.showinfo("Projection", "No data available to generate financial projection.")
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
            columns = [field['name'] for field in gui_fields] # Use GUI field names for display
            # Add primary key to columns for internal use, but don't show it as heading
            columns_with_pk = [primary_key] + columns

            tree = ttk.Treeview(frame, columns=columns_with_pk, show='headings')
            self.tabs[table_name]['tree'] = tree

            # Configure headings and columns
            for col_name in columns:
                tree.heading(col_name, text=col_name)
                tree.column(col_name, width=120, anchor='center')

            # Hide the primary key column
            tree.column(primary_key, width=0, stretch=tk.NO) # Hide it
            tree.heading(primary_key, text="") # No heading text

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
            columns = ['CategoryName']
            tree = ttk.Treeview(frame, columns=['CategoryID'] + columns, show='headings')
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
                # messagebox.showinfo("Load Data", f"No data available in {table_name}.") # Avoid too many popups
                return

            # Ensure correct column order for display based on gui_fields, and include PK for internal use
            display_columns = [field['name'] for field in TABLE_SCHEMAS[table_name]['gui_fields']]
            all_cols = [self.tabs[table_name]['primary_key']] + display_columns

            # Reorder DataFrame columns to match the Treeview's expected order
            # Also handle potential missing columns in DB by adding them as None
            for col in all_cols:
                if col not in df.columns:
                    df[col] = None # Add missing columns as None

            df = df[all_cols] # Reorder DataFrame columns

            for index, row in df.iterrows():
                values = tuple(row.values)
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

            fields_data = {} # To store Tkinter variables for each field
            entries = {}     # To store Tkinter widgets for each field

            selected_item_id = None
            current_record_data = {}
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
                pk_col = self.tabs[table_name]['primary_key']
                display_cols = [field['name'] for field in TABLE_SCHEMAS[table_name]['gui_fields']]
                all_cols_ordered = [pk_col] + display_cols

                # Create a dictionary from the values
                for i, col_name in enumerate(all_cols_ordered):
                    current_record_data[col_name] = values[i]

                logging.debug(f"Form: Editing record: {current_record_data}")


            row_num = 0
            for field_info in TABLE_SCHEMAS[table_name]['gui_fields']:
                field_name = field_info['name']
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
                    if field_info.get('allow_none', False):
                        combo_options.append("None") # Option for no selection

                    if 'options' in field_info:
                        combo_options.extend(field_info['options'])
                    elif 'source_table' in field_info:
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
                                    for idx, row in df_source.iterrows():
                                        # Safely access columns, provide fallback if column doesn't exist in df_source
                                        display_col = source_display_cols[i]
                                        value_col = source_value_cols[i]

                                        display_val = row[display_col] if display_col in row else f"MissingCol({display_col})"
                                        value_val = row[value_col] if value_col in row else f"MissingCol({value_col})"

                                        # Special handling for CategoryName where display and value are the same
                                        if display_col == value_col:
                                            combo_options.append(str(display_val))
                                        else:
                                            combo_options.append(f"{display_val} ({value_val})")
                            except Exception as e:
                                logging.error(f"Form: Error loading combo options from {src_tbl} for field {field_name}: {e}", exc_info=True)
                                # Continue to next source table or field if one fails
                    combo['values'] = combo_options
                    combo.set(combo_options[0] if combo_options else "") # Set default if available

                entries[field_name].grid(row=row_num, column=1, padx=5, pady=5, sticky='ew')
                row_num += 1

            # Populate fields if in edit mode
            if edit_mode:
                for field_info in TABLE_SCHEMAS[table_name]['gui_fields']:
                    field_name = field_info['name']
                    # The actual column name in the DB might differ from GUI field name (e.g., 'Category' vs 'CategoryID')
                    # Need to map GUI field name to actual DB column name for retrieval
                    db_col_name = None
                    if field_info['type'] == 'combo' and 'source_value_col' in field_info:
                        # If it's a combo box with a source table, the DB column is source_value_col
                        db_col_name = field_info['source_value_col']
                    else:
                        # Otherwise, the DB column name is usually the same as the GUI field name
                        db_col_name = field_name

                    if db_col_name and db_col_name in current_record_data:
                        value = current_record_data[db_col_name]
                        if entries[field_name].winfo_class() == 'TCombobox':
                            combo_widget = entries[field_name]
                            combo_values = combo_widget['values']
                            found_display_value = False

                            if field_info.get('allow_none') and (value is None or str(value).lower() == 'none' or value == 0): # Handle 0 for nullable IDs
                                combo_widget.set("None")
                                found_display_value = True
                            else:
                                # Try to find the display value from the original value (ID or Name)
                                for item_str in combo_values:
                                    # Match by value in parentheses for ID-based combos
                                    if f"({value})" in item_str and field_info['type'] == 'combo' and 'source_value_col' in field_info and field_info['source_value_col'] != field_info['source_display_col']:
                                        combo_widget.set(item_str)
                                        found_display_value = True
                                        break
                                    # Direct match for name-based combos (e.g., CategoryName, PaymentMethod)
                                    elif str(value) == item_str:
                                        combo_widget.set(item_str)
                                        found_display_value = True
                                        break

                            if not found_display_value and value is not None:
                                # Fallback: if value is not found, set the raw value (might not be pretty)
                                combo_widget.set(str(value))
                                logging.warning(f"Form: Could not find display value for {field_name} (DB col: {db_col_name}) with value {value}. Setting raw value.")
                        else:
                            entries[field_name].delete(0, tk.END)
                            entries[field_name].insert(0, value)
                    elif db_col_name not in current_record_data:
                        logging.warning(f"Form: Column '{db_col_name}' not found in current_record_data for {table_name}. Cannot pre-fill field {field_name}.")


            save_button = ttk.Button(form_window, text="Save", command=lambda: self.save_form_data(form_window, table_name, entries, edit_mode, current_record_data.get(pk_col)))
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

                # Determine the actual DB column name
                db_col_name = field_name
                if field_type == 'combo' and 'source_value_col' in field_info:
                    db_col_name = field_info['source_value_col']
                    # Extract the ID from the combo box string (e.g., "Creditor Name (ID)")
                    if '(' in value and ')' in value:
                        try:
                            value = value.split('(')[-1].replace(')', '')
                            if value.lower() == "none": # Handle explicit "None" selection
                                value = None
                            else:
                                value = int(value) # Convert ID to integer
                        except ValueError:
                            messagebox.showerror("Input Error", f"Invalid ID format for {field_name}. Please select a valid item or 'None'.")
                            logging.warning(f"SaveForm: Invalid ID format for {field_name}: {value}")
                            return
                    elif value.lower() == "none":
                        value = None # Explicitly set to None for 'None' selection
                    elif 'source_value_col' in field_info and field_info['source_value_col'] == field_info['source_display_col']:
                        # If display and value are the same (e.g., CategoryName), use value as is
                        pass
                    else:
                        messagebox.showerror("Input Error", f"Invalid selection for {field_name}. Please select from the list.")
                        logging.warning(f"SaveForm: Invalid selection for {field_name}: {value}")
                        return

                # Type conversion
                if value is None or value == '':
                    data_to_save[db_col_name] = None
                elif field_type == 'decimal':
                    try:
                        data_to_save[db_col_name] = float(value)
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid decimal value for {field_name}.")
                        logging.warning(f"SaveForm: Invalid decimal value for {field_name}: {value}")
                        return
                elif field_type == 'date':
                    try:
                        datetime.strptime(value, '%Y-%m-%d') # Validate format
                        data_to_save[db_col_name] = value # Store as string
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid date format for {field_name}. Use YYYY-MM-DD.")
                        logging.warning(f"SaveForm: Invalid date format for {field_name}: {value}")
                        return
                else:
                    data_to_save[db_col_name] = value

            success = False
            if edit_mode:
                success = db_manager.update_record(table_name, record_id, data_to_save)
                action = "updated"
            else:
                # For new records, handle 'Amount' and 'OriginalAmount' for Debts
                if table_name == 'Debts' and 'OriginalAmount' not in data_to_save and 'Amount' in data_to_save:
                    data_to_save['OriginalAmount'] = data_to_save['Amount']
                if table_name == 'Debts' and 'AmountPaid' not in data_to_save:
                    data_to_save['AmountPaid'] = 0.0 # Initialize AmountPaid for new debts

                success = db_manager.insert_record(table_name, data_to_save)
                action = "added"

            if success:
                messagebox.showinfo("Success", f"Record {action} successfully.")
                form_window.destroy()
                self.load_table_data(table_name) # Refresh the table data
                # Update related data if necessary
                if table_name == 'Payments' or table_name == 'Revenue':
                    db_manager.update_debt_amounts_and_payments()
                    self.load_dashboard_data() # Refresh dashboard for updated totals
                if table_name == 'Accounts':
                    db_manager.update_account_balances()
                    self.load_dashboard_data()
                if table_name == 'Goals':
                    db_manager.update_goal_progress()
                    self.load_dashboard_data()
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
                    if table_name == 'Payments' or table_name == 'Revenue':
                        db_manager.update_debt_amounts_and_payments()
                        self.load_dashboard_data()
                    if table_name == 'Accounts':
                        db_manager.update_account_balances()
                        self.load_dashboard_data()
                    if table_name == 'Goals':
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
                df_report = df_report[[col['name'] for col in TABLE_SCHEMAS['Debts']['columns'] if col['name'] != 'DebtID']] # Exclude ID
            elif report_type == 'Daily Expenses':
                try:
                    report_date = datetime.strptime(report_date_str, '%Y-%m-%d').strftime('%Y-%m-%d')
                except ValueError:
                    messagebox.showerror("Input Error", "Invalid date format. Please use YYYY-MM-DD.")
                    logging.warning(f"Reports: Invalid date format for Daily Expenses: {report_date_str}")
                    return
                payments_df = db_manager.get_table_data('Payments')
                df_report = payments_df[payments_df['PaymentDate'] == report_date]
                df_report = df_report[[col for col in df_report.columns if col not in ['PaymentID', 'DebtID']]] # Exclude IDs
            elif report_type == 'Snowball Progress':
                debts_df = db_manager.get_table_data('Debts')
                df_report = debts_df[
                    (debts_df['Status'] != 'Paid Off') &
                    (debts_df['Status'] != 'Closed')
                ].sort_values(by='Amount', ascending=True)
                df_report = df_report[['Creditor', 'Amount', 'MinimumPayment', 'SnowballPayment', 'Status']]
            elif report_type == 'Account Balances':
                df_report = db_manager.get_table_data('Accounts')
                df_report = df_report[[col for col in df_report.columns if col != 'AccountID']] # Exclude ID
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
            logging.info("AppClosing: User confirmed quit. Destroying application.")
            self.destroy()
        else:
            logging.info("AppClosing: User cancelled quit.")

if __name__ == "__main__":
    # Ensure database is initialized before launching the GUI
    try:
        logging.info("Main: Attempting to initialize database before GUI launch.")
        db_manager.initialize_database()
        logging.info("Main: Database initialized for GUI launch.")
        app = DebtManagerApp()
        app.mainloop()
    except Exception as e:
        logging.critical(f"Main: CRITICAL ERROR during GUI startup from __main__: {e}", exc_info=True)
        messagebox.showerror("Startup Error", f"Failed to initialize database or launch GUI: {e}\nCheck DebugLog.txt for more details.")

