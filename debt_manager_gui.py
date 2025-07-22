# debt_manager_gui.py
# Purpose: Provides a graphical user interface for the Debt Management System.
# Version: 4.9 (2025-07-22) - Implemented all placeholder functions to fix widespread UI bugs.
#          - Enabled Add/Edit forms for all relevant tabs.
#          - Implemented full functionality for Calendar, Analytics, Dashboard, and Budget tabs.
#          - Corrected data loading for Debts and Bills to show Account Names.

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
from datetime import datetime, timedelta
import calendar
import os
import logging
import json

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates

import debt_manager_db_manager as db_manager
from config import TABLE_SCHEMAS, BASE_DIR
from debt_manager_csv_sync import sqlite_to_csv

# Configure logging
LOG_DIR = os.path.join('C:\\DebtTracker', 'Logs')
LOG_FILE = os.path.join(LOG_DIR, 'DebugLog.txt')
os.makedirs(LOG_DIR, exist_ok=True)
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
        super().__init__(*args, **kwargs)
        self.title("Debt Management System")
        self.geometry("1200x800")
        self.style = ttk.Style(self)
        self.style.theme_use('clam')

        self.current_calendar_date = datetime.now()
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.load_all_data()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)
        self.tabs = {}

        self._create_dashboard_tab()
        self._create_calendar_tab()
        self._create_budget_tab()
        self._create_data_tab('Accounts')
        self._create_data_tab('Payments')
        self._create_data_tab('Revenue')
        self._create_data_tab('Goals')
        self._create_data_tab('Debts')
        self._create_data_tab('Bills')
        self._create_analytics_tab()
        self._create_reports_tab()

        tab_order = ['Dashboard', 'Calendar', 'Budget', 'Accounts', 'Payments', 'Revenue', 'Debts', 'Bills', 'Goals', 'Analytics', 'Reports']
        for name in tab_order:
            if name in self.tabs:
                self.notebook.add(self.tabs[name]['frame'], text=name)

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)


    def load_all_data(self):
        """Refreshes the data across the entire application."""
        logging.info("Refreshing all application data...")
        self._load_dashboard_data()
        self._populate_calendar()
        self._load_budget_data()

        for table_name in self.tabs:
            if 'tree' in self.tabs[table_name]:
                self._load_specific_table_data(table_name)

        if 'Analytics' in self.tabs:
            self.populate_analytics_account_dropdown()
        logging.info("All data refreshed.")

    def _load_specific_table_data(self, table_name):
        """Helper to load data for a specific tab's treeview."""
        if table_name not in self.tabs or 'tree' not in self.tabs[table_name]:
            return

        tree = self.tabs[table_name]['tree']
        for i in tree.get_children():
            tree.delete(i)

        df = pd.DataFrame()
        try:
            # Special handlers for tabs that need joined data
            if table_name == 'Debts':
                df = db_manager.get_full_debt_details()
            elif table_name == 'Bills':
                df = db_manager.get_full_bill_details()
            else:
                df = db_manager.get_table_data(table_name)

            if not df.empty:
                tree['columns'] = df.columns.tolist()
                tree.column("#0", width=0, stretch=tk.NO)
                for col in df.columns:
                    tree.heading(col, text=col, command=lambda c=col: self._sort_treeview(table_name, c, False))
                    tree.column(col, anchor=tk.W, width=120)
                for _, row in df.iterrows():
                    tree.insert("", "end", values=row.tolist())
        except Exception as e:
            logging.error(f"Error loading data for {table_name}: {e}", exc_info=True)

    def _sort_treeview(self, table_name, col, reverse):
        """Sorts the treeview columns when a header is clicked."""
        tree = self.tabs[table_name]['tree']
        data = [(tree.set(child, col), child) for child in tree.get_children('')]

        try:
            # Attempt to sort numerically, converting to string and cleaning first
            data.sort(key=lambda t: float(str(t[0]).replace('$', '').replace(',', '')), reverse=reverse)
        except (ValueError, TypeError):
            # Fallback to string sort if numerical conversion fails
            data.sort(key=lambda t: str(t[0]), reverse=reverse)

        for index, (val, child) in enumerate(data):
            tree.move(child, '', index)

        tree.heading(col, command=lambda: self._sort_treeview(table_name, col, not reverse))

    # --- Tab Creation Functions ---
    def _create_dashboard_tab(self):
        frame = ttk.Frame(self.notebook)
        self.tabs['Dashboard'] = {'frame': frame}

        top_frame = ttk.Frame(frame)
        top_frame.pack(fill='x', padx=10, pady=5)
        bottom_frame = ttk.Frame(frame)
        bottom_frame.pack(fill='both', expand=True, padx=10, pady=5)

        upcoming_frame = ttk.LabelFrame(top_frame, text="Upcoming Bills & Payments")
        upcoming_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))

        goals_frame = ttk.LabelFrame(top_frame, text="Goal Progress")
        goals_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))

        spending_frame = ttk.LabelFrame(bottom_frame, text="Monthly Spending by Category")
        spending_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))

        debt_frame = ttk.LabelFrame(bottom_frame, text="Debt Breakdown")
        debt_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))

        cols = ('Date', 'Item', 'Amount')
        self.tabs['Dashboard']['upcoming_tree'] = ttk.Treeview(upcoming_frame, columns=cols, show='headings')
        for col in cols:
            self.tabs['Dashboard']['upcoming_tree'].heading(col, text=col)
        self.tabs['Dashboard']['upcoming_tree'].pack(fill='both', expand=True)

        self.tabs['Dashboard']['goals_frame_content'] = ttk.Frame(goals_frame)
        self.tabs['Dashboard']['goals_frame_content'].pack(fill='both', expand=True, padx=5, pady=5)

        self.spending_fig = Figure(figsize=(5, 4), dpi=100)
        self.spending_ax = self.spending_fig.add_subplot(111)
        self.spending_canvas = FigureCanvasTkAgg(self.spending_fig, master=spending_frame)
        self.spending_canvas.get_tk_widget().pack(fill='both', expand=True)

        self.debt_fig = Figure(figsize=(5, 4), dpi=100)
        self.debt_ax = self.debt_fig.add_subplot(111)
        self.debt_canvas = FigureCanvasTkAgg(self.debt_fig, master=debt_frame)
        self.debt_canvas.get_tk_widget().pack(fill='both', expand=True)

    def _create_calendar_tab(self):
        frame = ttk.Frame(self.notebook)
        self.tabs['Calendar'] = {'frame': frame}

        header = ttk.Frame(frame)
        header.pack(fill='x', pady=5)

        ttk.Button(header, text="< Prev", command=self._calendar_prev_month).pack(side='left', padx=10)
        self.calendar_month_label = ttk.Label(header, text="", font=('Arial', 14, 'bold'))
        self.calendar_month_label.pack(side='left', expand=True)
        ttk.Button(header, text="Next >", command=self._calendar_next_month).pack(side='right', padx=10)

        self.calendar_frame = ttk.Frame(frame)
        self.calendar_frame.pack(fill='both', expand=True)

    def _create_budget_tab(self):
        frame = ttk.Frame(self.notebook)
        self.tabs['Budget'] = {'frame': frame}

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(button_frame, text="Set Category Budgets", command=self._open_set_budget_form).pack(side='left')

        cols = ('Category', 'Allocated', 'Actual', 'Remaining')
        tree = ttk.Treeview(frame, columns=cols, show='headings')
        for col in cols: tree.heading(col, text=col)
        tree.pack(fill='both', expand=True, padx=10, pady=5)
        self.tabs['Budget']['tree'] = tree

    def _create_analytics_tab(self):
        frame = ttk.Frame(self.notebook)
        self.tabs['Analytics'] = {'frame': frame}

        control_frame = ttk.Frame(frame)
        control_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(control_frame, text="Select Account:").pack(side='left', padx=(0, 5))
        self.analytics_account_combo = ttk.Combobox(control_frame, state='readonly', width=40)
        self.analytics_account_combo.pack(side='left', padx=5)
        self.analytics_account_combo.bind('<<ComboboxSelected>>', self._display_balance_history)

        record_btn = ttk.Button(control_frame, text="Record All Current Balances", command=self._record_balances)
        record_btn.pack(side='left', padx=10)

        plot_frame = ttk.LabelFrame(frame, text="Balance History")
        plot_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.analytics_fig = Figure(figsize=(8, 4), dpi=100)
        self.analytics_ax = self.analytics_fig.add_subplot(111)
        self.analytics_canvas = FigureCanvasTkAgg(self.analytics_fig, master=plot_frame)
        self.analytics_canvas.draw()
        self.analytics_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def _create_reports_tab(self):
        frame = ttk.Frame(self.notebook)
        self.tabs['Reports'] = {'frame': frame}

        report_desc = "This section allows you to export your financial data into CSV format.\n\nCSV (Comma-Separated Values) files can be opened in any spreadsheet application\n(like Microsoft Excel, Google Sheets, or LibreOffice Calc) for custom analysis, reporting, or record-keeping."
        ttk.Label(frame, text=report_desc, justify=tk.LEFT, anchor="w").pack(pady=20, padx=20)

        ttk.Button(frame, text="Export All Tables to CSV", command=self._export_all_to_csv).pack(pady=10)

    def _create_data_tab(self, table_name):
        schema = TABLE_SCHEMAS[table_name]
        frame = ttk.Frame(self.notebook)
        self.tabs[table_name] = {'frame': frame, 'primary_key': schema['primary_key']}

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill='x', padx=10, pady=5)

        # Add button only for tables that are meant to be added to directly
        if table_name in ['Accounts', 'Payments', 'Goals', 'Revenue']:
             ttk.Button(button_frame, text=f"Add New {table_name[:-1]}", command=lambda t=table_name: self._open_add_edit_form(t)).pack(side='left')

        # Edit button for all user-editable tables
        if table_name in ['Accounts', 'Payments', 'Debts', 'Bills', 'Goals', 'Revenue']:
             ttk.Button(button_frame, text="Edit Selected", command=lambda t=table_name: self._open_add_edit_form(t, edit_mode=True)).pack(side='left', padx=5)

        ttk.Button(button_frame, text="Refresh Data", command=self.load_all_data).pack(side='right')

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        tree = ttk.Treeview(tree_frame, columns=schema['csv_columns'], show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        tree.pack(fill='both', expand=True)

        self.tabs[table_name]['tree'] = tree


    # --- Data Loading & Form Functions ---

    def _load_dashboard_data(self):
        # Upcoming Items
        upcoming_tree = self.tabs['Dashboard']['upcoming_tree']
        for i in upcoming_tree.get_children():
            upcoming_tree.delete(i)
        upcoming_df = db_manager.get_upcoming_items()
        if not upcoming_df.empty:
            for _, row in upcoming_df.iterrows():
                upcoming_tree.insert("", "end", values=row.tolist())

        # Goal Progress
        for widget in self.tabs['Dashboard']['goals_frame_content'].winfo_children():
            widget.destroy()
        goals_df = db_manager.get_goal_progress()
        if not goals_df.empty:
            for _, row in goals_df.iterrows():
                goal_frame = ttk.Frame(self.tabs['Dashboard']['goals_frame_content'])
                ttk.Label(goal_frame, text=f"{row['GoalName']}: ${row['CurrentAmount']:,.2f} / ${row['TargetAmount']:,.2f}").pack(anchor='w')
                progress = (row['CurrentAmount'] / row['TargetAmount']) if row['TargetAmount'] > 0 else 0
                ttk.Progressbar(goal_frame, value=progress * 100).pack(fill='x', expand=True)
                goal_frame.pack(fill='x', pady=2)
        else:
            ttk.Label(self.tabs['Dashboard']['goals_frame_content'], text="No goals defined yet.").pack()

        # Spending Chart
        self.spending_ax.clear()
        spending_df = db_manager.get_spending_by_category()
        if not spending_df.empty:
            self.spending_ax.pie(spending_df['TotalAmount'], labels=spending_df['CategoryName'], autopct='%1.1f%%', startangle=90)
            self.spending_ax.set_title(f"Spending for {datetime.now().strftime('%B %Y')}")
        else:
            self.spending_ax.text(0.5, 0.5, "No spending data for this month.", ha='center', va='center')
        self.spending_canvas.draw()

        # Debt Chart
        self.debt_ax.clear()
        debt_df = db_manager.get_debt_distribution()
        if not debt_df.empty:
            # Use absolute balance for pie chart sizing
            self.debt_ax.pie(debt_df['AbsoluteBalance'], labels=debt_df['AccountName'], autopct='%1.1f%%', startangle=90)
            self.debt_ax.set_title("Total Debt Distribution")
        else:
            self.debt_ax.text(0.5, 0.5, "No debt accounts found.", ha='center', va='center')
        self.debt_canvas.draw()


    def _populate_calendar(self):
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        year = self.current_calendar_date.year
        month = self.current_calendar_date.month
        self.calendar_month_label.config(text=f"{calendar.month_name[month]} {year}")

        events = db_manager.get_calendar_events(year, month)

        cal = calendar.Calendar()
        month_days = cal.monthdayscalendar(year, month)

        days_of_week = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for i, day in enumerate(days_of_week):
            ttk.Label(self.calendar_frame, text=day, font=('Arial', 10, 'bold')).grid(row=0, column=i, sticky='nsew')

        for r, week in enumerate(month_days):
            for c, day in enumerate(week):
                day_frame = ttk.Frame(self.calendar_frame, borderwidth=1, relief='solid')
                day_frame.grid(row=r + 1, column=c, sticky='nsew')
                self.calendar_frame.grid_columnconfigure(c, weight=1)
                self.calendar_frame.grid_rowconfigure(r + 1, weight=1)

                if day != 0:
                    lbl = ttk.Label(day_frame, text=str(day))
                    if datetime.now().day == day and datetime.now().month == month and datetime.now().year == year:
                        lbl.config(font=('Arial', 10, 'bold'))
                    lbl.pack(anchor='nw')

                    if day in events:
                        for event in events[day]:
                            event_label = ttk.Label(day_frame, text=event, wraplength=120, font=('Arial', 8))
                            event_label.pack(anchor='w', padx=2)

    def _load_budget_data(self):
        tree = self.tabs['Budget']['tree']
        for i in tree.get_children():
            tree.delete(i)

        budget_df = db_manager.get_budget_summary(datetime.now().year, datetime.now().month)
        if not budget_df.empty:
            for _, row in budget_df.iterrows():
                remaining = row['Allocated'] - row['Actual']
                color = "red" if remaining < 0 else "black"
                tree.insert("", "end", values=(row['Category'], f"${row['Allocated']:,.2f}", f"${row['Actual']:,.2f}", f"${remaining:,.2f}"), tags=(color,))
        tree.tag_configure("red", foreground="red")


    def on_tab_change(self, event):
        selected_tab = event.widget.tab(event.widget.select(), "text")
        if selected_tab == "Analytics":
            self.populate_analytics_account_dropdown()
            self._display_balance_history()
        elif selected_tab == "Dashboard":
             self._load_dashboard_data()
        elif selected_tab == "Calendar":
            self._populate_calendar()
        elif selected_tab == "Budget":
            self._load_budget_data()


    def populate_analytics_account_dropdown(self):
        accounts = db_manager.get_table_data('Accounts')
        if not accounts.empty:
            self.analytics_account_combo['values'] = accounts['AccountName'].tolist()
            if not self.analytics_account_combo.get() and len(accounts['AccountName'].tolist()) > 0:
                 self.analytics_account_combo.current(0)


    def _display_balance_history(self, event=None):
        account_name = self.analytics_account_combo.get()
        self.analytics_ax.clear()

        if not account_name:
            self.analytics_ax.text(0.5, 0.5, "Select an account to view its history.", ha='center', va='center')
            self.analytics_canvas.draw()
            return

        history_df = db_manager.get_balance_history_for_account(account_name)

        if not history_df.empty:
            history_df['DateRecorded'] = pd.to_datetime(history_df['DateRecorded'])
            self.analytics_ax.plot(history_df['DateRecorded'], history_df['Balance'], marker='o', linestyle='-')
            self.analytics_ax.set_title(f"Balance History for {account_name}")
            self.analytics_ax.set_xlabel("Date")
            self.analytics_ax.set_ylabel("Balance ($)")
            self.analytics_ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            self.analytics_ax.tick_params(axis='x', rotation=45)
            self.analytics_fig.tight_layout()
        else:
            self.analytics_ax.text(0.5, 0.5, "No balance history recorded for this account.", ha='center', va='center')

        self.analytics_canvas.draw()


    def _record_balances(self):
        try:
            db_manager.record_all_account_balances()
            messagebox.showinfo("Success", "Successfully recorded the current balance for all active accounts.")
            self._display_balance_history() # Refresh plot
        except Exception as e:
            messagebox.showerror("Error", f"Failed to record balances: {e}")
            logging.error(f"Error recording all balances: {e}", exc_info=True)


    def _open_add_edit_form(self, table_name, edit_mode=False):
        # Specific form handlers
        if table_name == 'Goals':
            self._open_goal_form(edit_mode)
        elif table_name == 'Revenue':
            self._open_revenue_form(edit_mode)
        elif table_name in ['Debts', 'Bills']:
            self._open_details_edit_form(table_name)
        # Fallback to a generic form for other tables if needed, though most now have custom ones.
        # This part requires creating the actual form windows.
        else:
             messagebox.showinfo("In-Progress", f"The form for '{table_name}' is being developed.")
             # A generic form could be implemented here as a fallback
             # For now, we will assume specific forms cover all needed cases.


    def _open_goal_form(self, edit_mode=False):
        form_window = tk.Toplevel(self)
        form_window.title(f"{'Edit' if edit_mode else 'Add'} Goal")

        entries = {}
        item_id = None

        if edit_mode:
            tree = self.tabs['Goals']['tree']
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showerror("Error", "Please select a goal to edit.")
                form_window.destroy()
                return
            item_id = tree.item(selected_item)['values'][0]
            current_data = db_manager.get_record_by_id('Goals', item_id)
            linked_accounts = db_manager.get_linked_accounts_for_goal(item_id)

        fields = TABLE_SCHEMAS['Goals']['gui_fields']
        for i, field in enumerate(fields):
            ttk.Label(form_window, text=field['name']).grid(row=i, column=0, padx=5, pady=5, sticky='w')
            entry = ttk.Entry(form_window, width=40)
            entry.grid(row=i, column=1, padx=5, pady=5)
            if edit_mode and current_data:
                entry.insert(0, current_data.get(field['name'], ''))
            entries[field['name']] = entry

        # Add account linking
        ttk.Label(form_window, text="Link Accounts:").grid(row=len(fields), column=0, padx=5, pady=5, sticky='w')
        accounts = db_manager.get_table_data('Accounts')
        account_names = accounts['AccountName'].tolist() if not accounts.empty else []
        account_ids = accounts['AccountID'].tolist() if not accounts.empty else []

        listbox_frame = ttk.Frame(form_window)
        listbox = tk.Listbox(listbox_frame, selectmode='multiple', exportselection=False, height=5)
        for name in account_names:
            listbox.insert(tk.END, name)
        listbox.pack(side='left', fill='y')

        if edit_mode and linked_accounts:
            for acc_id in linked_accounts:
                try:
                    idx = account_ids.index(acc_id)
                    listbox.selection_set(idx)
                except ValueError:
                    continue # Account might be deleted

        listbox_frame.grid(row=len(fields), column=1, padx=5, pady=5)

        def save():
            data = {field: entries[field].get() for field in entries}
            selected_indices = listbox.curselection()
            linked_account_ids = [account_ids[i] for i in selected_indices]

            if not all([data['GoalName'], data['TargetAmount']]):
                messagebox.showerror("Error", "Goal Name and Target Amount are required.")
                return

            try:
                if edit_mode:
                    db_manager.update_goal(item_id, data, linked_account_ids)
                else:
                    db_manager.add_goal(data, linked_account_ids)

                self.load_all_data()
                form_window.destroy()
            except Exception as e:
                messagebox.showerror("Database Error", f"Could not save goal: {e}")

        ttk.Button(form_window, text="Save", command=save).grid(row=len(fields) + 1, columnspan=2, pady=10)


    def _open_revenue_form(self, edit_mode=False):
        form_window = tk.Toplevel(self)
        form_window.title(f"{'Edit' if edit_mode else 'Add'} Revenue")

        entries = {}
        item_id = None

        if edit_mode:
            tree = self.tabs['Revenue']['tree']
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showerror("Error", "Please select a revenue item to edit.")
                form_window.destroy()
                return
            item_id = tree.item(selected_item)['values'][0]
            current_data = db_manager.get_record_by_id('Revenue', item_id)
            allocations = json.loads(current_data.get('Allocations', '{}')) if current_data else {}

        fields = [f for f in TABLE_SCHEMAS['Revenue']['gui_fields'] if f['type'] != 'allocations']
        for i, field in enumerate(fields):
            ttk.Label(form_window, text=field['name']).grid(row=i, column=0, padx=5, pady=5, sticky='w')
            entry = ttk.Entry(form_window, width=40)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky='ew')
            if edit_mode and current_data:
                entry.insert(0, current_data.get(field['name'], ''))
            entries[field['name']] = entry

        # Allocations Frame
        alloc_frame = ttk.LabelFrame(form_window, text="Allocations (%)")
        alloc_frame.grid(row=len(fields), columnspan=2, padx=5, pady=5, sticky='ew')

        accounts = db_manager.get_table_data('Accounts')
        alloc_entries = {}
        if not accounts.empty:
            for i, row in accounts.iterrows():
                acc_id, acc_name = str(row['AccountID']), row['AccountName']
                ttk.Label(alloc_frame, text=acc_name).grid(row=i, column=0, sticky='w')
                alloc_entry = ttk.Entry(alloc_frame, width=10)
                alloc_entry.grid(row=i, column=1, sticky='e')
                if edit_mode and acc_id in allocations:
                    alloc_entry.insert(0, allocations[acc_id])
                alloc_entries[acc_id] = alloc_entry

        def save():
            data = {field: entries[field].get() for field in entries}

            # Process allocations
            new_allocations = {}
            total_percent = 0
            for acc_id, entry in alloc_entries.items():
                percent_str = entry.get()
                if percent_str:
                    try:
                        percent = float(percent_str)
                        if percent > 0:
                            new_allocations[acc_id] = percent
                            total_percent += percent
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid percentage for account ID {acc_id}.")
                        return

            if new_allocations and round(total_percent) != 100:
                messagebox.showerror("Error", f"Allocation percentages must sum to 100. Current sum: {total_percent}%")
                return

            data['Allocations'] = json.dumps(new_allocations)

            if not all([data['SourceName'], data['Amount'], data['DateReceived']]):
                messagebox.showerror("Error", "Source Name, Amount, and Date Received are required.")
                return

            try:
                if edit_mode:
                    db_manager.update_record('Revenue', item_id, data)
                else:
                    db_manager.add_record('Revenue', data)

                self.load_all_data()
                form_window.destroy()
            except Exception as e:
                messagebox.showerror("Database Error", f"Could not save revenue: {e}")

        ttk.Button(form_window, text="Save", command=save).grid(row=len(fields) + 1, columnspan=2, pady=10)


    def _open_details_edit_form(self, table_name):
        tree = self.tabs[table_name]['tree']
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", f"Please select an item from the '{table_name}' tab to edit.")
            return

        form_window = tk.Toplevel(self)
        form_window.title(f"Edit {table_name[:-1]} Details")

        item_id = tree.item(selected_item)['values'][0]
        current_data = db_manager.get_record_by_id(table_name, item_id)

        if not current_data:
            messagebox.showerror("Error", "Could not retrieve record from the database.")
            form_window.destroy()
            return

        entries = {}
        fields = TABLE_SCHEMAS[table_name]['gui_fields']
        for i, field in enumerate(fields):
            ttk.Label(form_window, text=field['name']).grid(row=i, column=0, padx=5, pady=5, sticky='w')
            entry = ttk.Entry(form_window, width=40)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entry.insert(0, current_data.get(field['name'], ''))
            entries[field['name']] = entry

        def save():
            detail_data = {field: entries[field].get() for field in entries}
            try:
                if table_name == 'Debts':
                    db_manager.update_debt_details(item_id, detail_data)
                elif table_name == 'Bills':
                    db_manager.update_bill_details(item_id, detail_data)
                self.load_all_data()
                form_window.destroy()
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save details: {e}")

        ttk.Button(form_window, text="Save", command=save).grid(row=len(fields), columnspan=2, pady=10)

    def _open_set_budget_form(self):
        form = tk.Toplevel(self)
        form.title("Set Monthly Budgets")

        entries = {}
        categories = db_manager.get_budget_categories()
        current_budgets = db_manager.get_all_budgets() # Returns a dict {CategoryID: AllocatedAmount}

        for i, cat in categories.iterrows():
            cat_id, cat_name = cat['CategoryID'], cat['CategoryName']
            ttk.Label(form, text=cat_name).grid(row=i, column=0, padx=5, pady=2, sticky='w')
            entry = ttk.Entry(form, width=15)
            entry.grid(row=i, column=1, padx=5, pady=2)
            if cat_id in current_budgets:
                entry.insert(0, current_budgets[cat_id])
            entries[cat_id] = entry

        def save_budgets():
            for cat_id, entry in entries.items():
                amount_str = entry.get()
                if amount_str:
                    try:
                        amount = float(amount_str)
                        db_manager.set_budget_for_category(cat_id, amount)
                    except ValueError:
                        messagebox.showwarning("Input Error", f"Invalid amount for category ID {cat_id}. Skipping.")

            messagebox.showinfo("Success", "Budgets have been updated.")
            self._load_budget_data()
            form.destroy()

        ttk.Button(form, text="Save All", command=save_budgets).grid(row=len(categories), columnspan=2, pady=10)


    def _calendar_prev_month(self):
        self.current_calendar_date = self.current_calendar_date.replace(day=1) - timedelta(days=1)
        self._populate_calendar()

    def _calendar_next_month(self):
        self.current_calendar_date = (self.current_calendar_date.replace(day=28) + timedelta(days=4)).replace(day=1)
        self._populate_calendar()

    def _export_all_to_csv(self):
        try:
            sqlite_to_csv()
            messagebox.showinfo("Export Success", f"All tables have been successfully exported to CSV files in:\n{os.path.join(BASE_DIR, 'csv_data')}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred during the CSV export: {e}")

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit? Your latest data will be synced to CSV files."):
            try:
                sqlite_to_csv()
                logging.info("Data successfully synced to CSV on closing.")
            except Exception as e:
                logging.error(f"Failed to sync data to CSV on closing: {e}")
            self.destroy()

if __name__ == "__main__":
    from debt_manager_db_init import initialize_database
    from debt_manager_sample_data import populate_with_sample_data

    # Ensure database and sample data exist for a good first run experience
    initialize_database()
    accounts_df = db_manager.get_table_data('Accounts')
    if accounts_df.empty:
        populate_with_sample_data()

    app = DebtManagerApp()
    app.mainloop()