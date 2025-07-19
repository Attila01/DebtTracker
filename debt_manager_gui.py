# debt_manager_gui.py
# Purpose: Main GUI application for the Debt Management System using Tkinter.
#          Handles all user interface logic, data operations, and interacts with
#          SQLite database and Excel synchronization functions.
# Deploy in: C:\DebtTracker
# Version: 2.10 (2025-07-18) - Fixed Credit Card tab data loading and available credit calculation.
#          Ensured AccountLimit is properly handled for Accounts and Debts.
#          Improved data type handling for numeric columns in get_table_data.

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import pandas as pd
import logging
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np # For numerical operations, especially for pie charts

# Import the synchronization and database initialization functions
from debt_manager_excel_sync import sync_data
from debt_manager_db_init import initialize_database # Although orchestrator will call it
from config import DB_PATH, EXCEL_PATH, REPORT_PATH, LOG_FILE, LOG_DIR, TABLE_SCHEMAS, PREDEFINED_CATEGORIES

# Ensure directories exist
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(REPORT_PATH, exist_ok=True)
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s: %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE, mode='a'),
                        logging.StreamHandler() # Also log to console
                    ])

# --- Database Operations ---
def get_db_connection():
    """Establishes and returns a SQLite database connection."""
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row # Allows accessing columns by name
        return conn
    except sqlite3.Error as e:
        logging.error(f"Database connection error: {e}")
        messagebox.showerror("Database Error", f"Could not connect to database: {e}")
        return None

def get_table_data(table_name):
    """
    Fetches all data from a specified table and ensures numeric columns are properly typed.
    """
    conn = get_db_connection()
    if conn is None:
        return pd.DataFrame(columns=TABLE_SCHEMAS[table_name]['db_columns']) # Return empty DataFrame with correct columns

    try:
        # Fetch all columns dynamically to avoid 'no such column' errors if schema is evolving
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)

        # Ensure all expected db_columns are present and in the correct order, filling missing with NaN
        expected_db_columns = TABLE_SCHEMAS[table_name]['db_columns']
        for col in expected_db_columns:
            if col not in df.columns:
                df[col] = None # Add missing columns as None initially

        df = df[expected_db_columns] # Reorder columns as per schema

        # Convert known numeric columns to float and fill NaNs with 0
        numeric_cols = [
            'Amount', 'OriginalAmount', 'AmountPaid', 'MinimumPayment', 'SnowballPayment',
            'InterestRate', 'Balance', 'StartingBalance', 'PreviousBalance', 'Value', 'TargetAmount',
            'CurrentAmount', 'AllocationPercentage', 'NextProjectedIncome', 'AccountLimit' # Added AccountLimit
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # Explicitly convert ID columns to integer, filling NaNs with 0
        id_cols = ['DebtID', 'AccountID', 'PaymentID', 'GoalID', 'AssetID', 'RevenueID', 'CategoryID', 'AllocatedTo']
        for col in id_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        logging.info(f"Loaded {table_name} ({len(df)} rows)")
        return df
    except pd.io.sql.DatabaseError as e:
        logging.warning(f"Table '{table_name}' not found or other database error when loading: {e}. Returning empty DataFrame.")
        return pd.DataFrame(columns=TABLE_SCHEMAS[table_name]['db_columns']) # Return empty with correct columns
    except Exception as e:
        logging.error(f"Load error ({table_name}): {e}")
        messagebox.showerror("Load Error", f"Could not load data from {table_name}: {e}")
        return pd.DataFrame(columns=TABLE_SCHEMAS[table_name]['db_columns'])
    finally:
        if conn:
            conn.close()

def update_table_data(table_name, df):
    """Updates the specified table in the database with the DataFrame content."""
    conn = get_db_connection()
    if conn is None:
        return False

    try:
        # Ensure the DataFrame columns match the DB schema before writing
        db_columns = TABLE_SCHEMAS[table_name]['db_columns']
        # Filter df to only include columns present in db_columns, and reorder them
        df_to_save = df[[col for col in db_columns if col in df.columns]]

        # Ensure any missing columns in df_to_save (that are in db_columns) are added as None
        for col in db_columns:
            if col not in df_to_save.columns:
                df_to_save[col] = None # Add as None, SQLite will handle NULL

        # Clear existing data and insert new data from DataFrame
        conn.execute(f"DELETE FROM {table_name}")
        df_to_save.to_sql(table_name, conn, if_exists='append', index=False)
        conn.commit()
        logging.info(f"Saved {table_name} ({len(df_to_save)} rows)")
        return True
    except sqlite3.Error as e:
        logging.error(f"Save error ({table_name}): {e}")
        messagebox.showerror("Save Error", f"Could not save changes to {table_name}: {e}")
        return False
    finally:
        if conn:
            conn.close()

def delete_record(table_name, primary_key_name, record_id):
    """Deletes a record from the specified table."""
    conn = get_db_connection()
    if conn is None:
        return False

    try:
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {table_name} WHERE {primary_key_name} = ?", (record_id,))
        conn.commit()
        logging.info(f"Deleted record {record_id} from {table_name}.")
        messagebox.showinfo("Delete Success", f"Record deleted from {table_name}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Delete error ({table_name}): {e}")
        messagebox.showerror("Delete Error", f"Could not delete record from {table_name}: {e}")
        return False
    finally:
        if conn:
            conn.close()

def add_record(table_name, data):
    """Adds a new record to the specified table."""
    conn = get_db_connection()
    if conn is None:
        return False

    try:
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?'] * len(data))
        values = tuple(data.values())

        cursor = conn.cursor()
        cursor.execute(f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})", values)
        conn.commit()
        logging.info(f"Added new record to {table_name}.")
        messagebox.showinfo("Add Success", f"New record added to {table_name}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Add record error ({table_name}): {e}")
        messagebox.showerror("Add Record Error", f"Could not add record to {table_name}: {e}")
        return False
    finally:
        if conn:
            conn.close()

def update_account_balances_and_debt_amounts():
    """
    Recalculates and updates account balances based on revenue and payments,
    and also updates debt 'Amount' and 'AmountPaid' based on payments.
    """
    conn = get_db_connection()
    if conn is None:
        return

    try:
        # Fetch data, ensuring numeric columns are already handled by get_table_data
        debts_df = get_table_data('Debts').copy()
        payments_df = get_table_data('Payments').copy()
        accounts_df = get_table_data('Accounts').copy()
        revenue_df = get_table_data('Revenue').copy()

        # --- Update Debt Amounts and AmountPaid ---
        if not debts_df.empty:
            # Calculate total payments made towards each debt
            # Ensure 'DebtID' in payments_df is integer for grouping
            payments_df['DebtID'] = pd.to_numeric(payments_df['DebtID'], errors='coerce').fillna(0).astype(int)
            debt_payments_sum = payments_df.groupby('DebtID')['Amount'].sum().reset_index()
            debt_payments_sum.rename(columns={'Amount': 'TotalPaid'}, inplace=True)

            # Merge payments sum with debts_df
            debts_df = pd.merge(debts_df, debt_payments_sum, on='DebtID', how='left')
            debts_df['TotalPaid'] = debts_df['TotalPaid'].fillna(0.0) # Ensure TotalPaid is float

            # Update AmountPaid and Amount
            debts_df['AmountPaid'] = debts_df['TotalPaid']
            debts_df['Amount'] = debts_df['OriginalAmount'] - debts_df['AmountPaid']
            debts_df['Amount'] = debts_df['Amount'].apply(lambda x: max(0.0, x)) # Ensure not negative

            # Update Status based on Amount
            debts_df['Status'] = debts_df.apply(lambda row: 'Paid Off' if row['Amount'] <= 0.01 else ('Paid' if row['Status'] == 'Paid' else 'Open'), axis=1)

            # Remove the temporary 'TotalPaid' column before saving
            debts_df = debts_df.drop(columns=['TotalPaid'])

            update_table_data('Debts', debts_df) # Save updated debts back to DB
            logging.info('Debt amounts and AmountPaid updated based on payments.')
        else:
            logging.info('No debts to update debt amounts.')

        # --- Update Account Balances ---
        if not accounts_df.empty:
            # Store current balances as previous balances before recalculating
            accounts_df['PreviousBalance'] = accounts_df['Balance']

            # Ensure AccountID in revenue_df and payments_df are numeric for matching
            revenue_df['AllocatedTo'] = pd.to_numeric(revenue_df['AllocatedTo'], errors='coerce').fillna(0).astype(int)
            payments_df['AccountID'] = pd.to_numeric(payments_df['AccountID'], errors='coerce').fillna(0).astype(int)

            for index, row in accounts_df.iterrows():
                account_id = row['AccountID']
                # StartingBalance is already ensured to be float and filled with 0.0 by get_table_data
                starting_balance = float(row['StartingBalance']) # Explicitly cast to float

                # Sum revenue allocated to this account
                deposits = revenue_df[(revenue_df['AllocatedTo'] == account_id) & (revenue_df['AllocationType'] == 'Account')]['Amount'].sum()
                deposits = float(deposits) # Ensure deposits is float

                # Sum payments made FROM this account
                withdrawals = payments_df[payments_df['AccountID'] == account_id]['Amount'].sum()
                withdrawals = float(withdrawals) # Ensure withdrawals is float

                new_balance = starting_balance + deposits - withdrawals
                accounts_df.at[index, 'Balance'] = new_balance

            update_table_data('Accounts', accounts_df) # Save updated accounts back to DB
            logging.info('Account balances updated.')
        else:
            logging.info('No accounts to update balances.')

    except sqlite3.Error as e:
        logging.error(f"Balance/Debt update error: {e}")
        messagebox.showerror("Update Error", f"Database error during balance/debt update: {e}")
    except Exception as e:
        logging.error(f"General update error in balances/debts: {e}")
        messagebox.showerror("Update Error", f"An unexpected error occurred during balance/debt update: {e}")
    finally:
        if conn:
            conn.close()

def generate_financial_projection(app_instance):
    """Generates a financial projection and exports it to a CSV file."""
    conn = get_db_connection()
    if conn is None:
        return

    try:
        cursor = conn.cursor()

        # Current Total Debt
        cursor.execute("SELECT SUM(Amount) FROM Debts WHERE Status NOT IN ('Paid Off', 'Closed')")
        total_debt = cursor.fetchone()[0] or 0.0

        # Current Total Savings (Accounts with 'Savings' or 'Checking' type)
        cursor.execute("SELECT SUM(Balance) FROM Accounts WHERE AccountType IN ('Savings', 'Checking', 'Investment') AND Status IN ('Open', 'Current', 'Active')")
        total_savings = cursor.fetchone()[0] or 0.0

        # Annual Income (sum revenue from last 12 months for 'current income')
        one_year_ago = (datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute("SELECT SUM(Amount) FROM Revenue WHERE DateReceived >= ?", (one_year_ago,))
        annual_income = cursor.fetchone()[0] or 0.0

        # Next Projected Income (sum from Revenue table)
        cursor.execute("SELECT SUM(NextProjectedIncome) FROM Revenue WHERE NextProjectedIncome IS NOT NULL")
        next_projected_income_sum = cursor.fetchone()[0] or 0.0

        # Debts for Snowball simulation (order by amount for smallest first)
        cursor.execute("SELECT DebtID, Creditor, Amount, MinimumPayment, SnowballPayment, InterestRate FROM Debts WHERE Status NOT IN ('Paid Off', 'Closed') ORDER BY Amount ASC")
        active_debts = [{'DebtID': row['DebtID'], 'Creditor': row['Creditor'], 'Amount': row['Amount'], 'MinPayment': row['MinimumPayment'], 'SnowballPayment': row['SnowballPayment'], 'InterestRate': row['InterestRate']} for row in cursor.fetchall()]

        projections = []

        # --- Simplified Projection Logic (Needs significant refinement for real-world accuracy) ---
        # This is a very basic simulation. For a true projection, a month-by-month
        # simulation considering interest, minimum payments, and actual snowballing would be needed.

        # Assumptions for simplified model:
        # - Net Monthly Income: (annual_income + next_projected_income_sum * 12) / 12 (approx monthly total)
        # - Fixed percentage of net income for debt payment and savings.
        # - Debt reduction is linear based on available payments.
        # - Savings grow at a simple rate.

        # Placeholder values for illustration
        monthly_income = (annual_income + next_projected_income_sum * 12) / 12 if annual_income > 0 or next_projected_income_sum > 0 else 2000 # Example if no data
        monthly_debt_payment_allocation_rate = 0.30 # 30% of income for debt
        monthly_savings_allocation_rate = 0.15 # 15% of income for savings
        monthly_expenses = monthly_income * (1 - monthly_debt_payment_allocation_rate - monthly_savings_allocation_rate) # Remaining for expenses

        current_debt = total_debt
        current_savings = total_savings

        projection_months = 12 * 10 # Project for 10 years

        monthly_projections = []

        for month in range(1, projection_months + 1):
            # Calculate available funds for debt/savings after expenses (simplified)
            available_for_debt_and_savings = monthly_income - monthly_expenses

            # Allocate to debt and savings
            payment_to_debt_this_month = min(current_debt, available_for_debt_and_savings * monthly_debt_payment_allocation_rate)
            contribution_to_savings_this_month = available_for_debt_and_savings * monthly_savings_allocation_rate

            current_debt -= payment_to_debt_this_month
            current_savings += contribution_to_savings_this_month

            # Simulate savings interest (simple monthly compound)
            current_savings *= (1 + (0.05 / 12)) # 5% annual interest

            # Record yearly snapshots
            if month % 12 == 0:
                year = month // 12
                projections.append({
                    'Year': year,
                    'ProjectedDebtRemaining': max(0.0, current_debt),
                    'ProjectedSavings': current_savings,
                    'ProjectedNetWorth': current_savings - current_debt
                })
                # If all debts are paid, allocate more to savings
                if current_debt <= 0.01:
                    monthly_savings_allocation_rate += monthly_debt_payment_allocation_rate # All debt allocation goes to savings
                    monthly_debt_payment_allocation_rate = 0.0
                    current_debt = 0.0

        df_projections = pd.DataFrame(projections)

        report_file = os.path.join(REPORT_PATH, f"FinancialProjection_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        df_projections.to_csv(report_file, index=False)
        logging.info(f"Projection generated: {report_file}")
        messagebox.showinfo("Projection Generated", f"Financial projection saved to:\n{report_file}")

        # Update the Goals graph
        app_instance.update_goals_graph(df_projections)


    except sqlite3.Error as e:
        logging.error(f"Projection error (SQLite): {e}")
        messagebox.showerror("Projection Error", f"Database error during projection: {e}")
    except Exception as e:
        logging.error(f"Projection error: {e}")
        messagebox.showerror("Projection Error", f"An unexpected error occurred during projection: {e}")
    finally:
        if conn:
            conn.close()


# --- GUI Class ---
class DebtManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Debt Management System")
        self.root.geometry("1200x750") # Increased size for graphs
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) # Handle window close event

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.data_frames = {} # To store pandas DataFrames for each tab
        self.category_map = self._get_category_map() # Load categories for dropdowns and filtering

        # IMPORTANT: Initialize reports tab FIRST, as it sets up self.fig_goals and self.ax_goals
        self._create_reports_tab()
        # Then create dashboard and other data tabs
        self._create_dashboard_tab()
        self._create_tabs()

        # Initial data load and sync on app launch
        # This initial refresh will now correctly trigger all updates and loads
        self._refresh_all_tabs()
        sync_data('sqlite_to_excel') # Push current DB state to Excel
        logging.info("Initial sync from SQLite to Excel completed on app launch.")

    def _get_category_map(self):
        """Fetches categories from DB and returns a map of {ID: Name, Name: ID}."""
        categories_df = get_table_data('Categories')
        category_id_to_name = {row['CategoryID']: row['CategoryName'] for idx, row in categories_df.iterrows()}
        category_name_to_id = {row['CategoryName']: row['CategoryID'] for idx, row in categories_df.iterrows()}
        return {'id_to_name': category_id_to_name, 'name_to_id': category_name_to_id}

    def _refresh_all_tabs(self):
        """Recalculates balances/debts, then refreshes data in all data tabs and dashboard."""
        # Ensure calculations are fresh before loading data into GUI
        update_account_balances_and_debt_amounts()

        # Refresh individual data tabs
        for table_name, schema in TABLE_SCHEMAS.items():
            tree = getattr(self, f"{table_name.lower()}_tree", None)
            if tree:
                # Special handling for Accounts tab to display sub-rows
                if table_name == 'Accounts':
                    self._load_accounts_with_allocations_to_treeview(tree)
                # Special handling for Debts tab to include Projected Payment
                elif table_name == 'Debts':
                    self._load_debts_to_treeview(tree)
                # Special handling for new category-based tabs
                elif table_name == 'Bills':
                    self._load_categorized_debts_to_treeview(tree, 'Bills')
                elif table_name == 'CreditCards':
                    self._load_categorized_debts_to_treeview(tree, 'Credit Card')
                elif table_name == 'Loans':
                    self._load_categorized_debts_to_treeview(tree, 'Loan')
                elif table_name == 'Collections':
                    self._load_categorized_debts_to_treeview(tree, 'Collection')
                else:
                    self._load_data_to_treeview(table_name, tree)

        # Refresh dashboard and its graphs
        self._load_dashboard_summary()
        self._create_dashboard_graphs()

        # Refresh goals graph (important for 'Goals' tab and 'Reports' tab)
        self.update_goals_graph()


    def _create_dashboard_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Dashboard")

        # Create a frame for the dashboard content
        dashboard_frame = ttk.Frame(tab)
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Dashboard Top Section - Summary Metrics
        summary_frame = ttk.LabelFrame(dashboard_frame, text="Summary Metrics")
        summary_frame.pack(fill="x", padx=10, pady=5)

        self.total_debt_label = ttk.Label(summary_frame, text="Total Outstanding Debt: $0.00", font=('Arial', 12, 'bold'))
        self.total_debt_label.pack(side="left", padx=10, pady=5)
        self.original_total_debt_label = ttk.Label(summary_frame, text="Original Total Debt: $0.00", font=('Arial', 12, 'bold'))
        self.original_total_debt_label.pack(side="left", padx=10, pady=5)
        self.active_debts_label = ttk.Label(summary_frame, text="Number of Active Debts: 0", font=('Arial', 12, 'bold'))
        self.active_debts_label.pack(side="left", padx=10, pady=5)
        self.total_paid_label = ttk.Label(summary_frame, text="Total Paid on Debts: $0.00", font=('Arial', 12, 'bold'))
        self.total_paid_label.pack(side="left", padx=10, pady=5)
        self.total_savings_label = ttk.Label(summary_frame, text="Total Savings/Assets: $0.00", font=('Arial', 12, 'bold'))
        self.total_savings_label.pack(side="left", padx=10, pady=5)
        self.net_worth_label = ttk.Label(summary_frame, text="Net Worth: $0.00", font=('Arial', 12, 'bold'))
        self.net_worth_label.pack(side="left", padx=10, pady=5)
        self.current_income_label = ttk.Label(summary_frame, text="Last 30 Days Income: $0.00", font=('Arial', 12, 'bold'))
        self.current_income_label.pack(side="left", padx=10, pady=5)
        self.next_projected_income_label = ttk.Label(summary_frame, text="Next Projected Monthly Income: $0.00", font=('Arial', 12, 'bold'))
        self.next_projected_income_label.pack(side="left", padx=10, pady=5)
        self.next_payment_due_label = ttk.Label(summary_frame, text="Next Payment Due: N/A", font=('Arial', 12, 'bold'))
        self.next_payment_due_label.pack(side="left", padx=10, pady=5)


        # Alerts & Reminders Section
        alerts_frame = ttk.LabelFrame(dashboard_frame, text="Alerts & Reminders")
        alerts_frame.pack(fill="x", padx=10, pady=5)
        self.alerts_text = tk.Text(alerts_frame, height=3, wrap="word", state="disabled")
        self.alerts_text.pack(fill="both", expand=True, padx=5, pady=5)


        # Snowball Data Treeview
        snowball_frame = ttk.LabelFrame(dashboard_frame, text="Debt Snowball/Avalanche Order")
        snowball_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.dashboard_tree = ttk.Treeview(snowball_frame, show="headings")
        self.dashboard_tree.pack(fill="both", expand=True, side="left", padx=(0,0), pady=(0,0)) # Adjusted padding

        # Scrollbar for Treeview
        scrollbar = ttk.Scrollbar(snowball_frame, orient="vertical", command=self.dashboard_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.dashboard_tree.configure(yscrollcommand=scrollbar.set)

        # Dashboard Buttons
        button_frame = ttk.Frame(dashboard_frame)
        button_frame.pack(fill="x", padx=10, pady=5)

        load_snowball_button = ttk.Button(button_frame, text="Refresh Dashboard Data", command=self._load_dashboard_summary)
        load_snowball_button.pack(side="left", padx=5, pady=5)

        # Frame for Dashboard Graphs
        self.dashboard_graphs_frame = ttk.LabelFrame(dashboard_frame, text="Financial Visuals")
        self.dashboard_graphs_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Create figure and axes for dashboard graphs (will be cleared and redrawn)
        self.dashboard_fig, (self.ax_debt_dist, self.ax_cash_flow) = plt.subplots(1, 2, figsize=(10, 4))
        self.dashboard_canvas = FigureCanvasTkAgg(self.dashboard_fig, master=self.dashboard_graphs_frame)
        self.dashboard_canvas_widget = self.dashboard_canvas.get_tk_widget()
        self.dashboard_canvas_widget.pack(fill="both", expand=True)


    def _load_dashboard_summary(self):
        """Loads and displays summary metrics and snowball data on the dashboard."""
        conn = get_db_connection()
        if conn is None:
            return

        try:
            cursor = conn.cursor()

            # --- Summary Metrics ---
            debts_df = get_table_data('Debts')
            accounts_df = get_table_data('Accounts')
            assets_df = get_table_data('Assets')
            revenue_df = get_table_data('Revenue')
            payments_df = get_table_data('Payments')

            total_debt = debts_df[debts_df['Status'].isin(['Open', 'Paid'])]['Amount'].sum()
            original_total_debt = debts_df['OriginalAmount'].sum()
            active_debts_count = debts_df[debts_df['Status'].isin(['Open', 'Paid'])].shape[0]
            total_paid_on_debts = debts_df['AmountPaid'].sum()

            total_savings = accounts_df[accounts_df['AccountType'].isin(['Savings', 'Checking', 'Investment'])]['Balance'].sum()
            total_assets = assets_df[assets_df['AssetStatus'] == 'Active']['Value'].sum()
            net_worth = (total_savings + total_assets) - total_debt

            self.total_debt_label.config(text=f"Total Outstanding Debt: ${total_debt:,.2f}")
            self.original_total_debt_label.config(text=f"Original Total Debt: ${original_total_debt:,.2f}")
            self.active_debts_label.config(text=f"Number of Active Debts: {active_debts_count}")
            self.total_paid_label.config(text=f"Total Paid on Debts: ${total_paid_on_debts:,.2f}")
            self.total_savings_label.config(text=f"Total Savings/Assets: ${total_savings + total_assets:,.2f}")
            self.net_worth_label.config(text=f"Net Worth: ${(total_savings + total_assets) - total_debt:,.2f}")

            # Current Income (last 30 days)
            thirty_days_ago = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d %H:%M:%S')
            current_income = revenue_df[revenue_df['DateReceived'] >= thirty_days_ago]['Amount'].sum()
            self.current_income_label.config(text=f"Last 30 Days Income: ${current_income:,.2f}")

            # Next Projected Income
            next_projected_income = revenue_df['NextProjectedIncome'].sum() if 'NextProjectedIncome' in revenue_df.columns else 0.0
            self.next_projected_income_label.config(text=f"Next Projected Monthly Income: ${next_projected_income:,.2f}")

            # Next Payment Due
            # Ensure 'PaymentDate' is converted to datetime objects for proper comparison and sorting
            payments_df['PaymentDate_dt'] = pd.to_datetime(payments_df['PaymentDate'], errors='coerce')
            upcoming_payments = payments_df[
                (payments_df['PaymentDate_dt'] >= datetime.now()) # Use datetime object for comparison
            ].sort_values(by='PaymentDate_dt')

            next_payment_info = "N/A"
            if not upcoming_payments.empty:
                next_payment = upcoming_payments.iloc[0]
                next_payment_date = next_payment['PaymentDate_dt'].strftime('%Y-%m-%d') # Format from datetime object
                next_payment_info = f"${next_payment['Amount']:,.2f} on {next_payment_date} (DebtID: {next_payment['DebtID']})"
            self.next_payment_due_label.config(text=f"Next Payment Due: {next_payment_info}")

            # --- Alerts & Reminders ---
            alerts_messages = []
            # Low account balance warnings (example: below $100)
            low_balance_accounts = accounts_df[accounts_df['Balance'] < 100]
            for idx, acc in low_balance_accounts.iterrows():
                alerts_messages.append(f"WARNING: Low balance in {acc['AccountName']} (${acc['Balance']:,.2f})")

            # Upcoming due dates (next 7 days)
            seven_days_from_now = (datetime.now() + timedelta(days=7))
            # Ensure 'DueDate' is converted to datetime objects for proper comparison
            debts_df['DueDate_dt'] = pd.to_datetime(debts_df['DueDate'], errors='coerce')
            upcoming_debt_dues = debts_df[
                (debts_df['DueDate_dt'] >= datetime.now()) &
                (debts_df['DueDate_dt'] <= seven_days_from_now) &
                (debts_df['Status'].isin(['Open', 'Paid']))
            ].sort_values(by='DueDate_dt')

            for idx, debt in upcoming_debt_dues.iterrows():
                alerts_messages.append(f"REMINDER: {debt['Creditor']} payment of ${debt['MinimumPayment']:,.2f} due by {debt['DueDate_dt'].strftime('%Y-%m-%d')}")

            self.alerts_text.config(state="normal")
            self.alerts_text.delete(1.0, tk.END)
            if alerts_messages:
                self.alerts_text.insert(tk.END, "\n".join(alerts_messages))
            else:
                self.alerts_text.insert(tk.END, "No new alerts or reminders.")
            self.alerts_text.config(state="disabled")


            # --- Snowball Data ---
            # Order by Amount (smallest first for Snowball) or InterestRate (highest first for Avalanche)
            # Let's default to Snowball (smallest amount first)
            active_debts_for_snowball = debts_df[debts_df['Status'].isin(['Open', 'Paid'])].sort_values(by='Amount', ascending=True)

            # Include Projected Payment in dashboard view
            columns = ['DebtID', 'Creditor', 'Amount', 'OriginalAmount', 'AmountPaid', 'MinimumPayment', 'SnowballPayment', 'InterestRate', 'DueDate', 'Status']
            display_columns = ['Debt ID', 'Creditor', 'Current Amount', 'Original Amount', 'Amount Paid', 'Min Payment', 'Snowball Payment', 'Interest Rate', 'Due Date', 'Status']

            # Clear existing data in Treeview
            self.dashboard_tree.delete(*self.dashboard_tree.get_children())

            # Set up columns if not already set (or refresh them)
            self.dashboard_tree["columns"] = columns
            for i, col in enumerate(columns):
                self.dashboard_tree.heading(col, text=display_columns[i])
                if 'ID' in col:
                    self.dashboard_tree.column(col, width=50, anchor="center")
                elif 'Amount' in col or 'Payment' in col or 'InterestRate' in col:
                    self.dashboard_tree.column(col, width=100, anchor="e") # Right align numbers
                elif 'Date' in col:
                    self.dashboard_tree.column(col, width=100, anchor="center")
                else:
                    self.dashboard_tree.column(col, width=120, anchor="w")

            # Insert data
            for index, row in active_debts_for_snowball.iterrows():
                values = [row[col] for col in columns]
                # Format currency and date values for display
                values[2] = f"${values[2]:,.2f}" if isinstance(values[2], (int, float)) else values[2] # Current Amount
                values[3] = f"${values[3]:,.2f}" if isinstance(values[3], (int, float)) else values[3] # Original Amount
                values[4] = f"${values[4]:,.2f}" if isinstance(values[4], (int, float)) else values[4] # Amount Paid
                values[5] = f"${values[5]:,.2f}" if isinstance(values[5], (int, float)) else values[5] # Minimum Payment
                values[6] = f"${values[6]:,.2f}" if isinstance(values[6], (int, float)) else values[6] # Snowball Payment
                values[7] = f"{values[7]:.2f}%" if isinstance(values[7], (int, float)) else values[7] # Interest Rate
                # Format DueDate from datetime object if it was converted, or from string if not
                if isinstance(row['DueDate_dt'], datetime):
                    values[8] = row['DueDate_dt'].strftime('%Y-%m-%d')
                elif isinstance(values[8], str) and ' ' in values[8]: # Fallback for string
                    try:
                        values[8] = datetime.strptime(values[8], '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                    except ValueError:
                        pass # Keep as is if format doesn't match

                self.dashboard_tree.insert("", "end", values=values)

            logging.info(f"Dashboard summary and snowball data loaded ({len(active_debts_for_snowball)} debts).")

        except sqlite3.Error as e:
            logging.error(f"Dashboard load error: {e}")
            messagebox.showerror("Load Error", f"Could not load dashboard data: {e}")
        except Exception as e:
            logging.error(f"Dashboard general error: {e}")
            messagebox.showerror("Load Error", f"An unexpected error occurred while loading dashboard: {e}")
        finally:
            if conn:
                conn.close()

    def _create_dashboard_graphs(self):
        """Creates and updates the graphs on the dashboard."""
        self.ax_debt_dist.clear()
        self.ax_cash_flow.clear()

        debts_df = get_table_data('Debts')
        revenue_df = get_table_data('Revenue')
        payments_df = get_table_data('Payments')

        # --- Debt Distribution Pie Chart ---
        active_debts = debts_df[debts_df['Status'].isin(['Open', 'Paid'])]
        if not active_debts.empty:
            debt_amounts = active_debts['Amount']
            debt_labels = active_debts['Creditor']

            # Filter out zero amounts for pie chart
            non_zero_debts = debt_amounts[debt_amounts > 0]
            non_zero_labels = debt_labels[debt_amounts > 0]

            if not non_zero_debts.empty:
                self.ax_debt_dist.pie(non_zero_debts, labels=non_zero_labels, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 8})
                self.ax_debt_dist.set_title('Debt Distribution by Creditor', fontsize=10)
                self.ax_debt_dist.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.
            else:
                self.ax_debt_dist.text(0.5, 0.5, "No active debts for distribution.", horizontalalignment='center', verticalalignment='center', transform=self.ax_debt_dist.transAxes, fontsize=8)
        else:
            self.ax_debt_dist.text(0.5, 0.5, "No active debts for distribution.", horizontalalignment='center', verticalalignment='center', transform=self.ax_debt_dist.transAxes, fontsize=8)

        # --- Cash Flow Trend Line Chart ---
        # Combine revenue and payments, convert dates to datetime objects
        revenue_df['Date'] = pd.to_datetime(revenue_df['DateReceived'], errors='coerce')
        payments_df['Date'] = pd.to_datetime(payments_df['PaymentDate'], errors='coerce')

        # Filter out invalid dates
        revenue_df = revenue_df.dropna(subset=['Date'])
        payments_df = payments_df.dropna(subset=['Date'])

        # Create monthly cash flow
        monthly_income = revenue_df.set_index('Date').resample('M')['Amount'].sum().fillna(0)
        monthly_expenses = payments_df.set_index('Date').resample('M')['Amount'].sum().fillna(0)

        # Align indices (dates)
        all_months = monthly_income.index.union(monthly_expenses.index)
        monthly_income = monthly_income.reindex(all_months, fill_value=0)
        monthly_expenses = monthly_expenses.reindex(all_months, fill_value=0)

        monthly_net_cash_flow = monthly_income - monthly_expenses

        if not monthly_net_cash_flow.empty:
            self.ax_cash_flow.plot(monthly_net_cash_flow.index, monthly_net_cash_flow.values, marker='o', linestyle='-', color='purple')
            self.ax_cash_flow.set_title('Monthly Net Cash Flow Trend', fontsize=10)
            self.ax_cash_flow.set_xlabel('Date', fontsize=8)
            self.ax_cash_flow.set_ylabel('Net Cash Flow ($)', fontsize=8)
            self.ax_cash_flow.grid(True, linestyle='--', alpha=0.6)
            self.ax_cash_flow.tick_params(axis='x', rotation=45, labelsize=7)
            self.ax_cash_flow.tick_params(axis='y', labelsize=7)
            self.ax_cash_flow.axhline(0, color='grey', linestyle='--', linewidth=0.8) # Zero line
        else:
            self.ax_cash_flow.text(0.5, 0.5, "No cash flow data for trend.", horizontalalignment='center', verticalalignment='center', transform=self.ax_cash_flow.transAxes, fontsize=8)

        self.dashboard_fig.tight_layout()
        self.dashboard_canvas.draw()


    def _create_tabs(self):
        # Create standard data tabs
        for table_name, schema in TABLE_SCHEMAS.items():
            # Skip dashboard and reports tabs as they are created separately
            # Also skip Bills, CreditCards, Loans, Collections as they are derived tabs
            if table_name in ['Dashboard', 'Reports', 'Bills', 'CreditCards', 'Loans', 'Collections']:
                continue

            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=table_name)

            tree_frame = ttk.Frame(tab)
            tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

            tree = ttk.Treeview(tree_frame, show="headings")
            tree.pack(fill="both", expand=True, side="left")

            # Use db_columns for internal Treeview column names, excel_columns for display headers
            # Special handling for Debts to add 'Projected Payment'
            if table_name == 'Debts':
                debt_cols = TABLE_SCHEMAS['Debts']['db_columns'] + ['ProjectedPayment']
                debt_display_cols = TABLE_SCHEMAS['Debts']['excel_columns'] + ['Projected Payment']
                self._setup_treeview_columns(tree, debt_cols, debt_display_cols)
            else:
                self._setup_treeview_columns(tree, schema['db_columns'], schema['excel_columns'])

            scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            scrollbar.pack(side="right", fill="y")
            tree.configure(yscrollcommand=scrollbar.set)

            setattr(self, f"{table_name.lower()}_tree", tree)

            button_frame = ttk.Frame(tab)
            button_frame.pack(fill="x", padx=10, pady=5)

            refresh_button = ttk.Button(button_frame, text="Refresh Data", command=lambda: self._refresh_all_tabs())
            refresh_button.pack(side="left", padx=5, pady=5)

            add_button = ttk.Button(button_frame, text=f"Add {table_name}", command=lambda tn=table_name, s=schema['fields_for_new'], t=tree: self._open_add_record_form(tn, s, t))
            add_button.pack(side="left", padx=5, pady=5)

            edit_button = ttk.Button(button_frame, text="Edit Selected", command=lambda tn=table_name, pk=schema['primary_key'], s=schema['fields_for_new'], t=tree: self._open_edit_record_form(tn, pk, s, t))
            edit_button.pack(side="left", padx=5, pady=5)

            delete_button = ttk.Button(button_frame, text="Delete Selected", command=lambda tn=table_name, pk=schema['primary_key'], t=tree: self._delete_selected_record(tn, pk, t))
            delete_button.pack(side="left", padx=5, pady=5)

            sync_to_excel_button = ttk.Button(button_frame, text="Sync DB to Excel", command=lambda: self._trigger_excel_sync('sqlite_to_excel'))
            sync_to_excel_button.pack(side="left", padx=5, pady=5)

            sync_from_excel_button = ttk.Button(button_frame, text="Sync Excel to DB", command=lambda: self._trigger_excel_sync('excel_to_sqlite'))
            sync_from_excel_button.pack(side="left", padx=5, pady=5)

            if table_name == 'Goals':
                update_goal_button = ttk.Button(button_frame, text="Update Goal Progress", command=self._update_goal_progress)
                update_goal_button.pack(side="left", padx=5, pady=5)
                generate_proj_button = ttk.Button(button_frame, text="Generate Financial Projection", command=lambda: generate_financial_projection(self))
                generate_proj_button.pack(side="left", padx=5, pady=5)
                self.update_goals_graph()

        # --- Create New Specialized Tabs (derived from Debts) ---
        self._create_derived_debt_tab('Bills', 'Bills', ['Debt ID', 'Creditor', 'Monthly Bill', 'Amount Paid', 'Remaining Balance', 'Forecast'])
        self._create_derived_debt_tab('Credit Cards', 'Credit Card', ['Debt ID', 'Creditor', 'Current Balance', 'Available Credit', 'Linked Account ID'])
        self._create_derived_debt_tab('Loans', 'Loan', ['Debt ID', 'Creditor', 'Current Balance', 'Available Credit', 'Linked Account ID'])
        self._create_derived_debt_tab('Collections', 'Collection', ['Debt ID', 'Creditor', 'Current Balance', 'Status', 'Notes'])


    def _create_derived_debt_tab(self, tab_name, category_name, display_columns):
        """Helper to create new category-based debt tabs."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=tab_name)

        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tree = ttk.Treeview(tree_frame, show="headings")
        tree.pack(fill="both", expand=True, side="left")

        # Set up columns for derived tabs
        # Internal column names will be simple indexed names like 'col1', 'col2'
        # as the data is dynamically calculated/filtered.
        tree_cols = [f'col{i+1}' for i in range(len(display_columns))]
        self._setup_treeview_columns(tree, tree_cols, display_columns)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)

        setattr(self, f"{tab_name.lower().replace(' ', '')}_tree", tree) # Store with no spaces

        button_frame = ttk.Frame(tab)
        button_frame.pack(fill="x", padx=10, pady=5)

        refresh_button = ttk.Button(button_frame, text="Refresh Data", command=lambda: self._refresh_all_tabs())
        refresh_button.pack(side="left", padx=5, pady=5)


    def _setup_treeview_columns(self, tree, db_columns, display_columns):
        tree["columns"] = db_columns # Internal column names for Treeview
        for i, col_name in enumerate(db_columns):
            display_text = display_columns[i] if i < len(display_columns) else col_name # Fallback
            tree.heading(col_name, text=display_text)
            if 'ID' in col_name or 'Category ID' in display_text:
                tree.column(col_name, width=80, anchor="center")
            elif 'Amount' in col_name or 'Balance' in col_name or 'Payment' in col_name or 'Value' in col_name or 'Income' in col_name or 'Credit' in col_name or 'Monthly Bill' in display_text or 'Forecast' in display_text:
                tree.column(col_name, width=120, anchor="e") # Right align numbers
            elif 'Date' in col_name:
                tree.column(col_name, width=100, anchor="center")
            else:
                tree.column(col_name, width=150, anchor="w")

    def _load_data_to_treeview(self, table_name, tree):
        """Loads data for standard tables."""
        df = get_table_data(table_name)
        self.data_frames[table_name] = df # Store DataFrame for potential updates

        # Clear existing data in Treeview
        tree.delete(*tree.get_children())

        # Insert new data
        for index, row in df.iterrows():
            # Ensure the row values match the tree's internal column order (db_columns)
            ordered_values = []
            for col in TABLE_SCHEMAS[table_name]['db_columns']:
                value = row[col]
                # Format currency and date values for display in treeview
                if isinstance(value, (int, float)) and ('Amount' in col or 'Balance' in col or 'Payment' in col or 'Value' in col or 'Income' in col or 'Limit' in col):
                    ordered_values.append(f"${value:,.2f}")
                elif isinstance(value, (int, float)) and 'InterestRate' in col:
                    ordered_values.append(f"{value:.2f}%")
                elif col == 'CategoryID': # Display Category Name instead of ID
                    category_name = self.category_map['id_to_name'].get(value, "Unknown")
                    ordered_values.append(category_name)
                elif isinstance(value, str) and ('Date' in col or 'Received' in col) and ' ' in value: # Check for datetime string
                    try:
                        ordered_values.append(datetime.strptime(value, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d'))
                    except ValueError:
                        ordered_values.append(value) # Keep original if format doesn't match
                else:
                    ordered_values.append(value)

            tree.insert("", "end", values=ordered_values)
        logging.info(f"GUI: Loaded data for '{table_name}' tab.")

    def _load_debts_to_treeview(self, tree):
        """Loads data for Debts tab, including Projected Payment."""
        debts_df = get_table_data('Debts')
        self.data_frames['Debts'] = debts_df # Store DataFrame

        # Clear existing data in Treeview
        tree.delete(*tree.get_children())

        # Insert new data
        for index, row in debts_df.iterrows():
            # Calculate Projected Payment (simplified: MinPayment + SnowballPayment)
            projected_payment = row['MinimumPayment'] + row['SnowballPayment']

            ordered_values = []
            for col in TABLE_SCHEMAS['Debts']['db_columns']:
                value = row[col]
                if isinstance(value, (int, float)) and ('Amount' in col or 'Payment' in col or 'InterestRate' in col):
                    ordered_values.append(f"${value:,.2f}")
                elif col == 'CategoryID':
                    category_name = self.category_map['id_to_name'].get(value, "Unknown")
                    ordered_values.append(category_name)
                elif isinstance(value, str) and 'Date' in col and ' ' in value:
                    try:
                        ordered_values.append(datetime.strptime(value, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d'))
                    except ValueError:
                        ordered_values.append(value)
                else:
                    ordered_values.append(value)

            ordered_values.append(f"${projected_payment:,.2f}") # Add Projected Payment

            tree.insert("", "end", values=ordered_values)
        logging.info(f"GUI: Loaded data for 'Debts' tab with Projected Payments.")

    def _load_accounts_with_allocations_to_treeview(self, tree):
        """Loads data for Accounts tab, showing allocations as sub-rows."""
        accounts_df = get_table_data('Accounts')
        payments_df = get_table_data('Payments')
        revenue_df = get_table_data('Revenue')
        self.data_frames['Accounts'] = accounts_df # Store DataFrame

        tree.delete(*tree.get_children())

        for index, account_row in accounts_df.iterrows():
            account_id = account_row['AccountID']

            # Main Account Row
            account_values = []
            for col in TABLE_SCHEMAS['Accounts']['db_columns']:
                value = account_row[col]
                if isinstance(value, (int, float)) and ('Balance' in col or 'Limit' in col):
                    account_values.append(f"${value:,.2f}")
                else:
                    account_values.append(value)

            parent_iid = tree.insert("", "end", values=account_values, tags=('account_row',))

            # Add Payments as sub-rows
            account_payments = payments_df[payments_df['AccountID'] == account_id]
            for p_idx, payment_row in account_payments.iterrows():
                payment_values = [
                    "", # Indent for sub-row
                    f"Payment: {payment_row['PaymentMethod']}",
                    "", # No balance
                    f"-${payment_row['Amount']:,.2f}", # Negative for outflow
                    "", "", "", "", # Fill empty columns
                    payment_row['PaymentDate'].split(' ')[0], # Date
                    payment_row['Category'],
                    payment_row['Notes']
                ]
                tree.insert(parent_iid, "end", values=payment_values, tags=('payment_subrow',))

            # Add Revenue as sub-rows
            account_revenue = revenue_df[(revenue_df['AllocatedTo'] == account_id) & (revenue_df['AllocationType'] == 'Account')]
            for r_idx, revenue_row in account_revenue.iterrows():
                revenue_values = [
                    "", # Indent for sub-row
                    f"Deposit: {revenue_row['Source']}",
                    "", # No balance
                    f"+${revenue_row['Amount']:,.2f}", # Positive for inflow
                    "", "", "", "", # Fill empty columns
                    revenue_row['DateReceived'].split(' ')[0], # Date
                    revenue_row['AllocationType'],
                    "" # No notes for revenue currently
                ]
                tree.insert(parent_iid, "end", values=revenue_values, tags=('revenue_subrow',))

        # Apply tags for styling (optional, but good for visual distinction)
        tree.tag_configure('account_row', font=('Arial', 10, 'bold'))
        tree.tag_configure('payment_subrow', foreground='red')
        tree.tag_configure('revenue_subrow', foreground='green')

        logging.info(f"GUI: Loaded data for 'Accounts' tab with allocations.")


    def _load_categorized_debts_to_treeview(self, tree, category_name):
        """Loads and calculates data for Bills, Credit Cards, Loans, Collections tabs."""
        debts_df = get_table_data('Debts')
        accounts_df = get_table_data('Accounts')

        category_id = self.category_map['name_to_id'].get(category_name)
        if category_id is None:
            logging.warning(f"Category '{category_name}' not found for tab.")
            tree.delete(*tree.get_children())
            tree.insert("", "end", values=[f"No debts found for category '{category_name}'."], tags=('no_data',))
            return

        filtered_debts = debts_df[debts_df['CategoryID'] == category_id]

        tree.delete(*tree.get_children())

        if filtered_debts.empty:
            tree.insert("", "end", values=[f"No debts found for category '{category_name}'."], tags=('no_data',))
            return

        for index, row in filtered_debts.iterrows():
            debt_id = row['DebtID']
            creditor = row['Creditor']
            current_amount = row['Amount']
            amount_paid = row['AmountPaid']
            status = row['Status']
            linked_account_id = row['AccountID'] # This is the AccountID from Debts table

            display_values = []

            if category_name == 'Bills':
                monthly_bill = row['MinimumPayment']
                remaining_balance = current_amount
                forecast = remaining_balance + monthly_bill # Simple forecast for next bill
                display_values = [
                    debt_id,
                    creditor,
                    f"${monthly_bill:,.2f}",
                    f"${amount_paid:,.2f}",
                    f"${remaining_balance:,.2f}",
                    f"${forecast:,.2f}"
                ]
            elif category_name == 'Credit Card':
                account_limit = 0.0
                # Find the linked account in the Accounts DataFrame
                if linked_account_id and linked_account_id != 0:
                    linked_account = accounts_df[accounts_df['AccountID'] == linked_account_id]
                    if not linked_account.empty:
                        account_limit = linked_account.iloc[0].get('AccountLimit', 0.0) # Use .get with default for safety

                available_credit = max(0.0, account_limit - current_amount) # Available credit = Limit - Current Balance
                display_values = [
                    debt_id,
                    creditor,
                    f"${current_amount:,.2f}",
                    f"${available_credit:,.2f}",
                    linked_account_id if linked_account_id != 0 else "N/A"
                ]
            elif category_name == 'Loan': # Loans tab uses same logic as Credit Cards for now
                account_limit = 0.0
                if linked_account_id and linked_account_id != 0:
                    linked_account = accounts_df[accounts_df['AccountID'] == linked_account_id]
                    if not linked_account.empty:
                        account_limit = linked_account.iloc[0].get('AccountLimit', 0.0)

                available_credit = max(0.0, account_limit - current_amount)
                display_values = [
                    debt_id,
                    creditor,
                    f"${current_amount:,.2f}",
                    f"${available_credit:,.2f}",
                    linked_account_id if linked_account_id != 0 else "N/A"
                ]
            elif category_name == 'Collection':
                notes = row.get('Notes', '') # Assuming Notes might be in Debts table
                display_values = [
                    debt_id,
                    creditor,
                    f"${current_amount:,.2f}",
                    status,
                    notes
                ]

            tree.insert("", "end", values=display_values)

        logging.info(f"GUI: Loaded data for '{category_name}' tab.")


    def _refresh_tab_data(self, table_name, tree):
        """Refreshes data for a specific tab from the database."""
        # Now calls the centralized refresh method
        self._refresh_all_tabs()
        messagebox.showinfo("Data Refreshed", f"Data for '{table_name}' tab refreshed from database.")

    def _trigger_excel_sync(self, sync_direction):
        """Triggers the Excel sync process and refreshes all GUI data."""
        logging.info(f"GUI: Triggering Excel sync: {sync_direction}")
        try:
            sync_data(sync_direction)
            messagebox.showinfo("Sync Complete", f"Data synchronized {sync_direction.replace('_', ' ')}. Refreshing all data.")
            # Now calls the centralized refresh method
            self._refresh_all_tabs()
        except Exception as e:
            messagebox.showerror("Sync Error", f"Error during Excel sync ({sync_direction}): {e}")
            logging.error(f"GUI: Error during Excel sync ({sync_direction}): {e}")

    def _delete_selected_record(self, table_name, primary_key_name, tree):
        selected_item = tree.focus()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a record to delete.")
            return

        values = tree.item(selected_item, 'values')
        if not values:
            messagebox.showwarning("No Data", "Selected row has no data.")
            return

        # Assuming primary key is the first column in the Treeview values
        record_id = values[0]

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete record {record_id} from {table_name}?"):
            if delete_record(table_name, primary_key_name, record_id):
                # Now calls the centralized refresh method
                self._refresh_all_tabs()
                messagebox.showinfo("Refresh & Sync", "Record deleted. Data refreshed and synced to Excel.")

    def _open_add_record_form(self, table_name, fields_schema, parent_tree):
        add_form = tk.Toplevel(self.root)
        add_form.title(f"Add New {table_name} Record")
        add_form.transient(self.root) # Make it appear on top of the main window
        add_form.grab_set() # Disable interaction with main window until this is closed

        entries = {}
        row_num = 0

        # Helper to get current debt/account names for dropdowns
        def get_debt_account_options():
            options = ["0 - None"]
            conn = get_db_connection()
            if conn:
                cursor = conn.cursor()
                cursor.execute("SELECT DebtID, Creditor FROM Debts")
                options.extend([f"{row['DebtID']} - {row['Creditor']} (Debt)" for row in cursor.fetchall()])
                cursor.execute("SELECT AccountID, AccountName FROM Accounts")
                options.extend([f"{row['AccountID']} - {row['AccountName']} (Account)" for row in cursor.fetchall()])
                conn.close()
            return options

        def get_account_options():
            options = ["0 - None"]
            conn = get_db_connection()
            if conn:
                cursor = conn.cursor()
                cursor.execute("SELECT AccountID, AccountName FROM Accounts")
                options.extend([f"{row['AccountID']} - {row['AccountName']}" for row in cursor.fetchall()])
                conn.close()
            return options

        def get_category_options():
            # Use predefined categories from config for dropdown
            return [cat['CategoryName'] for cat in PREDEFINED_CATEGORIES]

        for field in fields_schema:
            label_text = field.get('label', field['name']) # Use 'label' if provided, else 'name'
            label = ttk.Label(add_form, text=f"{label_text}:")
            label.grid(row=row_num, column=0, padx=5, pady=5, sticky="w")

            if field['type'] == 'dropdown':
                combo = ttk.Combobox(add_form, values=field['options'], state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                if field['options']:
                    combo.set(field['options'][0]) # Set default
                entries[field['name']] = combo
            elif field['type'] == 'date':
                entry = ttk.Entry(add_form)
                entry.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                entry.insert(0, datetime.now().strftime('%Y-%m-%d')) # Default to current date
                entries[field['name']] = entry
            elif field['type'] == 'debt_account_selector': # For Payments (DebtID), Revenue (AllocatedTo)
                combo_values = get_debt_account_options()
                combo = ttk.Combobox(add_form, values=combo_values, state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                combo.set(combo_values[0])
                entries[field['name']] = combo
            elif field['type'] == 'account_selector': # For Payments (AccountID), Debts (AccountID)
                combo_values = get_account_options()
                combo = ttk.Combobox(add_form, values=combo_values, state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                combo.set(combo_values[0])
                entries[field['name']] = combo
            elif field['type'] == 'category_dropdown':
                combo_values = get_category_options()
                combo = ttk.Combobox(add_form, values=combo_values, state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                if combo_values:
                    combo.set(combo_values[0])
                entries[field['name']] = combo
            else: # Default to text entry for 'text' and 'real' (will validate later)
                entry = ttk.Entry(add_form)
                entry.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                entries[field['name']] = entry

            row_num += 1

        def save_new_record():
            new_data = {}
            for field in fields_schema:
                value = entries[field['name']].get().strip()
                if field['type'] == 'real':
                    try:
                        new_data[field['name']] = float(value) if value else 0.0
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid number for {field['name']}.")
                        return
                elif field['type'] == 'date':
                    try:
                        # Store dates as 'YYYY-MM-DD HH:MM:SS' for SQLite TEXT
                        new_data[field['name']] = datetime.strptime(value, '%Y-%m-%d').strftime('%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid date format for {field['name']}. Use YYYY-MM-DD.")
                        return
                elif field['type'] in ['debt_account_selector', 'account_selector']:
                    # Extract ID from "ID - Name" string, handle "0 - None"
                    if " - " in value:
                        try:
                            new_data[field['name']] = int(value.split(' - ')[0])
                        except ValueError:
                            new_data[field['name']] = None # Or handle as 0 for "None"
                    else:
                        new_data[field['name']] = None # No selection / invalid
                elif field['type'] == 'category_dropdown':
                    # Convert category name back to ID for storage
                    new_data[field['name']] = self.category_map['name_to_id'].get(value)
                else: # text, dropdown
                    new_data[field['name']] = value if value else None # Store empty string as None/NULL

            # Handle default values for new columns not in add form (e.g., 'Balance' for Accounts, 'Amount' for Debts in some cases)
            if table_name == 'Accounts':
                if 'Balance' not in new_data:
                    new_data['Balance'] = new_data.get('StartingBalance', 0.0)
                if 'StartingBalance' not in new_data: # Ensure StartingBalance is set if not provided
                    new_data['StartingBalance'] = new_data.get('Balance', 0.0)
                if 'PreviousBalance' not in new_data: # Set previous balance for new accounts
                    new_data['PreviousBalance'] = new_data.get('StartingBalance', 0.0)
                if 'AccountLimit' not in new_data:
                    new_data['AccountLimit'] = 0.0 # Default for new accounts

            if table_name == 'Debts':
                if 'OriginalAmount' not in new_data and 'Amount' in new_data:
                    new_data['OriginalAmount'] = new_data['Amount']
                elif 'Amount' not in new_data and 'OriginalAmount' in new_data:
                    new_data['Amount'] = new_data['OriginalAmount'] # If only original provided, current is same
                elif 'Amount' not in new_data and 'OriginalAmount' not in new_data:
                    new_data['OriginalAmount'] = 0.0
                    new_data['Amount'] = 0.0 # Default if neither provided

                if 'AmountPaid' not in new_data:
                    new_data['AmountPaid'] = 0.0 # Default to 0 for new debts
                if 'CategoryID' not in new_data:
                    new_data['CategoryID'] = None # Default to None if not selected
                if 'AccountID' not in new_data:
                    new_data['AccountID'] = None # Default to None if not selected

            if table_name == 'Revenue':
                if 'AllocationPercentage' not in new_data:
                    new_data['AllocationPercentage'] = 0.0
                if 'NextProjectedIncome' not in new_data:
                    new_data['NextProjectedIncome'] = 0.0
                if 'NextProjectedIncomeDate' not in new_data:
                    new_data['NextProjectedIncomeDate'] = None # Or default to empty string
                if new_data.get('AllocationType') == 'Category-Based':
                    new_data['AllocatedTo'] = None # Clear AllocatedTo if category-based

            if table_name == 'Payments':
                if 'Notes' not in new_data:
                    new_data['Notes'] = None # Default to None if not provided
                if 'AccountID' not in new_data:
                    new_data['AccountID'] = None # Default to None if not provided

            if table_name == 'Assets':
                if 'Category' not in new_data:
                    new_data['Category'] = None # Default to None if not provided

            if add_record(table_name, new_data):
                add_form.destroy()
                # Now calls the centralized refresh method
                self._refresh_all_tabs()
                messagebox.showinfo("Refresh & Sync", "New record added. Data refreshed and synced to Excel.")

        save_button = ttk.Button(add_form, text="Save", command=save_new_record)
        save_button.grid(row=row_num, column=0, columnspan=2, pady=10)

        add_form.update_idletasks() # Update window to get actual size
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (add_form.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (add_form.winfo_height() // 2)
        add_form.geometry(f"+{x}+{y}")

    def _open_edit_record_form(self, table_name, primary_key_name, fields_schema, parent_tree):
        selected_item = parent_tree.focus()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a record to edit.")
            return

        selected_values = parent_tree.item(selected_item, 'values')
        if not selected_values:
            messagebox.showwarning("No Data", "Selected row has no data to edit.")
            return

        # Get the primary key value and ensure it's an integer for lookup
        try:
            record_id_to_edit = int(selected_values[0])
        except ValueError:
            messagebox.showerror("Error", "Invalid record ID format. Cannot edit.")
            logging.error(f"Invalid record ID for {table_name}: {selected_values[0]}")
            return

        current_df = self.data_frames.get(table_name)
        if current_df is None or current_df.empty:
            messagebox.showerror("Error", f"No data loaded for {table_name} to edit.")
            return

        pk_col_name = TABLE_SCHEMAS[table_name]['primary_key']

        # Ensure the primary key column in the DataFrame is also integer type for accurate comparison
        # This is a safer way to ensure int type for PK column if it somehow became float during data loading
        if pd.api.types.is_numeric_dtype(current_df[pk_col_name]) and not pd.api.types.is_integer_dtype(current_df[pk_col_name]):
            current_df[pk_col_name] = current_df[pk_col_name].astype(int)

        # Find the row in the DataFrame by primary key
        # Use .copy() to avoid SettingWithCopyWarning later if modifying this slice
        record_data_series = current_df[current_df[pk_col_name] == record_id_to_edit]

        if record_data_series.empty:
            messagebox.showerror("Error", f"Record with ID {record_id_to_edit} not found in {table_name} data.")
            logging.error(f"Record ID {record_id_to_edit} not found in DataFrame for {table_name}.")
            return

        record_data = record_data_series.iloc[0].to_dict()

        edit_form = tk.Toplevel(self.root)
        edit_form.title(f"Edit {table_name} Record (ID: {record_id_to_edit})")
        edit_form.transient(self.root)
        edit_form.grab_set()

        entries = {}
        row_num = 0

        # Helper to get current debt/account names for dropdowns
        def get_debt_account_options():
            options = ["0 - None"]
            conn = get_db_connection()
            if conn:
                cursor = conn.cursor()
                cursor.execute("SELECT DebtID, Creditor FROM Debts")
                options.extend([f"{row['DebtID']} - {row['Creditor']} (Debt)" for row in cursor.fetchall()])
                cursor.execute("SELECT AccountID, AccountName FROM Accounts")
                options.extend([f"{row['AccountID']} - {row['AccountName']} (Account)" for row in cursor.fetchall()])
                conn.close()
            return options

        def get_account_options():
            options = ["0 - None"]
            conn = get_db_connection()
            if conn:
                cursor = conn.cursor()
                cursor.execute("SELECT AccountID, AccountName FROM Accounts")
                options.extend([f"{row['AccountID']} - {row['AccountName']}" for row in cursor.fetchall()])
                conn.close()
            return options

        def get_category_options():
            return [cat['CategoryName'] for cat in PREDEFINED_CATEGORIES]

        for field in fields_schema:
            label_text = field.get('label', field['name'])
            label = ttk.Label(edit_form, text=f"{label_text}:")
            label.grid(row=row_num, column=0, padx=5, pady=5, sticky="w")

            current_value = record_data.get(field['name'])
            # Ensure current_value is not None for display or conversion attempts
            if current_value is None:
                current_value = ""

            if field['type'] == 'dropdown':
                combo = ttk.Combobox(edit_form, values=field['options'], state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                combo.set(str(current_value))
                entries[field['name']] = combo
            elif field['type'] == 'date':
                entry = ttk.Entry(edit_form)
                entry.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                # Format date from DB ('YYYY-MM-DD HH:MM:SS') to 'YYYY-MM-DD' for display
                if isinstance(current_value, str) and ' ' in current_value:
                    try:
                        # Attempt to parse as datetime, then format to YYYY-MM-DD
                        entry.insert(0, datetime.strptime(current_value, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d'))
                    except ValueError:
                        entry.insert(0, current_value) # Fallback if format doesn't match
                else:
                    entry.insert(0, str(current_value))
                entries[field['name']] = entry
            elif field['type'] in ['debt_account_selector', 'account_selector']:
                combo_values = get_debt_account_options() if field['type'] == 'debt_account_selector' else get_account_options()
                combo = ttk.Combobox(edit_form, values=combo_values, state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")

                selected_display_value = "0 - None"
                if current_value is not None and current_value != "": # Check for empty string too
                    try:
                        # Convert current_value to int for comparison with option IDs
                        current_id_int = int(current_value)
                        for opt in combo_values:
                            if opt.startswith(f"{current_id_int} - "):
                                selected_display_value = opt
                                break
                    except ValueError:
                        pass # If current_value is not a valid integer ID, keep default
                combo.set(selected_display_value)
                entries[field['name']] = combo
            elif field['type'] == 'category_dropdown':
                combo_values = get_category_options()
                combo = ttk.Combobox(edit_form, values=combo_values, state="readonly")
                combo.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")

                # Set selected value based on current CategoryID
                current_category_name = self.category_map['id_to_name'].get(current_value, "")
                combo.set(current_category_name)
                entries[field['name']] = combo
            else: # text, real
                entry = ttk.Entry(edit_form)
                entry.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
                entry.insert(0, str(current_value))
                entries[field['name']] = entry

            row_num += 1

        # --- Special handling for Debts table: Display Current Amount (Calculated) ---
        if table_name == 'Debts':
            current_amount_label = ttk.Label(edit_form, text="Current Amount (Calculated):")
            current_amount_label.grid(row=row_num, column=0, padx=5, pady=5, sticky="w")

            current_amount_value = record_data.get('Amount', 0.0)
            current_amount_display = ttk.Entry(edit_form, state='readonly') # Make it read-only
            current_amount_display.grid(row=row_num, column=1, padx=5, pady=5, sticky="ew")
            current_amount_display.insert(0, f"${current_amount_value:,.2f}")
            row_num += 1


        def save_edited_record():
            updated_data = record_data.copy()
            for field in fields_schema:
                value = entries[field['name']].get().strip()
                if field['type'] == 'real':
                    try:
                        updated_data[field['name']] = float(value) if value else 0.0
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid number for {field['name']}.")
                        return
                elif field['type'] == 'date':
                    try:
                        updated_data[field['name']] = datetime.strptime(value, '%Y-%m-%d').strftime('%Y-%m-%d %H:%M:%S')
                    except ValueError:
                        messagebox.showerror("Input Error", f"Invalid date format for {field['name']}. Use YYYY-MM-DD.")
                        return
                elif field['type'] in ['debt_account_selector', 'account_selector']:
                    if " - " in value:
                        try:
                            updated_data[field['name']] = int(value.split(' - ')[0])
                        except ValueError:
                            updated_data[field['name']] = None
                    else:
                        updated_data[field['name']] = None
                elif field['type'] == 'category_dropdown':
                    updated_data[field['name']] = self.category_map['name_to_id'].get(value)
                else:
                    updated_data[field['name']] = value if value else None

            # Find the row to update in the DataFrame using the original integer ID
            # Use .loc for label-based indexing, which is safer with integer IDs
            idx_to_update = self.data_frames[table_name].index[self.data_frames[table_name][primary_key_name] == record_id_to_edit].tolist()

            if idx_to_update: # If list is not empty, means record was found
                # Log state before update for debugging
                logging.info(f"DEBUG: Before update for {table_name} ID {record_id_to_edit}:")
                logging.info(f"  Current DataFrame row: {self.data_frames[table_name].loc[idx_to_update[0]].to_dict()}")
                logging.info(f"  Updated data from form: {updated_data}")

                # Update the row in the DataFrame
                for col, val in updated_data.items():
                    if col in TABLE_SCHEMAS[table_name]['db_columns']:
                        self.data_frames[table_name].loc[idx_to_update[0], col] = val # Use idx_to_update[0] as it's a single index

                if update_table_data(table_name, self.data_frames[table_name]):
                    edit_form.destroy()
                    # Now calls the centralized refresh method
                    self._refresh_all_tabs()
                    messagebox.showinfo("Update Success", f"Record {record_id_to_edit} in {table_name} updated successfully. Data refreshed and synced to Excel.")

                    # Log state after full refresh for debugging
                    # Re-fetch the data for the specific record to ensure it's truly updated
                    refreshed_df = get_table_data(table_name)
                    refreshed_record = refreshed_df[refreshed_df[primary_key_name] == record_id_to_edit]
                    if not refreshed_record.empty:
                        logging.info(f"DEBUG: After full refresh for {table_name} ID {record_id_to_edit}:")
                        logging.info(f"  Refreshed OriginalAmount: {refreshed_record.iloc[0].get('OriginalAmount', 'N/A'):.2f}")
                        logging.info(f"  Refreshed AmountPaid: {refreshed_record.iloc[0].get('AmountPaid', 'N/A'):.2f}")
                        logging.info(f"  Refreshed Calculated Amount: {refreshed_record.iloc[0].get('Amount', 'N/A'):.2f}")
                        logging.info(f"  Refreshed CategoryID: {refreshed_record.iloc[0].get('CategoryID', 'N/A')}")
                        logging.info(f"  Refreshed AccountID: {refreshed_record.iloc[0].get('AccountID', 'N/A')}")
                    else:
                        logging.warning(f"DEBUG: Could not find record {record_id_to_edit} after refresh.")

                else:
                    messagebox.showerror("Update Failed", "Failed to save updated record to database.")
            else:
                messagebox.showerror("Error", "Could not find record to update in DataFrame.")

        save_button = ttk.Button(edit_form, text="Save Changes", command=save_edited_record)
        save_button.grid(row=row_num, column=0, columnspan=2, pady=10)

        edit_form.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (edit_form.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (edit_form.winfo_height() // 2)
        edit_form.geometry(f"+{x}+{y}")

    def _create_reports_tab(self):
        """Creates the Reports tab with options to generate reports."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Reports")

        report_frame = ttk.LabelFrame(tab, text="Generate Reports")
        report_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Financial Projection Report
        proj_button = ttk.Button(report_frame, text="Generate Financial Projection (CSV)", command=lambda: generate_financial_projection(self))
        proj_button.pack(pady=10)

        # Future reports can be added here
        # E.g., Debt Summary Report, Account Balances Report etc.

        # Goals Progress Graph (will be updated by generate_financial_projection)
        self.goals_graph_frame = ttk.LabelFrame(report_frame, text="Goals Progress Projection")
        self.goals_graph_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.fig_goals, self.ax_goals = plt.subplots(figsize=(8, 4))
        self.canvas_goals = FigureCanvasTkAgg(self.fig_goals, master=self.goals_graph_frame)
        self.canvas_goals_widget = self.canvas_goals.get_tk_widget()
        self.canvas_goals_widget.pack(fill="both", expand=True)
        self.update_goals_graph() # Initial empty graph

    def update_goals_graph(self, projection_df=None):
        """
        Updates the Goals Progress graph.
        Can take an optional projection_df to plot projected net worth.
        """
        self.ax_goals.clear()

        goals_df = get_table_data('Goals')

        if not goals_df.empty:
            # Plot actual goal progress
            for index, row in goals_df.iterrows():
                goal_name = row['GoalName']
                target_amount = row['TargetAmount']
                current_amount = row['CurrentAmount']

                # Simple bar for current progress vs target
                self.ax_goals.barh(goal_name, current_amount, color='skyblue', label='Current Progress')
                self.ax_goals.barh(goal_name, target_amount - current_amount, left=current_amount, color='lightgrey', label='Remaining')

                self.ax_goals.text(current_amount, goal_name, f'${current_amount:,.2f}', va='center', ha='left', fontsize=8)
                self.ax_goals.text(target_amount, goal_name, f'Target: ${target_amount:,.2f}', va='center', ha='right', fontsize=8)

            self.ax_goals.set_title('Goals Progress', fontsize=10)
            self.ax_goals.set_xlabel('Amount ($)', fontsize=8)
            self.ax_goals.tick_params(axis='x', labelsize=7)
            self.ax_goals.tick_params(axis='y', labelsize=8)
            self.ax_goals.legend(["Current Progress", "Remaining"], loc='upper right', fontsize=7)
            self.ax_goals.grid(axis='x', linestyle='--', alpha=0.7)
        else:
            self.ax_goals.text(0.5, 0.5, "No goals defined.", horizontalalignment='center', verticalalignment='center', transform=self.ax_goals.transAxes, fontsize=8)

        # Plot projected net worth if projection_df is provided
        if projection_df is not None and not projection_df.empty:
            # Create a secondary axis for net worth if needed, or overlay on a different plot type
            # For simplicity, let's just plot it as a line on the same axes if it fits, or clear and plot only projection
            # If we overlay, ensure scales make sense.

            # Clear previous goal bars if we are showing a new projection
            self.ax_goals.clear()
            self.ax_goals.plot(projection_df['Year'], projection_df['ProjectedNetWorth'], marker='o', linestyle='-', color='green', label='Projected Net Worth')
            self.ax_goals.set_title('Projected Net Worth Over Time', fontsize=10)
            self.ax_goals.set_xlabel('Year', fontsize=8)
            self.ax_goals.set_ylabel('Net Worth ($)', fontsize=8)
            self.ax_goals.grid(True, linestyle='--', alpha=0.6)
            self.ax_goals.tick_params(axis='x', labelsize=7)
            self.ax_goals.tick_params(axis='y', labelsize=7)
            self.ax_goals.axhline(0, color='grey', linestyle='--', linewidth=0.8) # Zero line
            self.ax_goals.legend(loc='upper left', fontsize=7)


        self.fig_goals.tight_layout()
        self.canvas_goals.draw()

    def _update_goal_progress(self):
        """Opens a form to update the current amount of a selected goal."""
        goals_df = get_table_data('Goals')
        if goals_df.empty:
            messagebox.showinfo("No Goals", "Please add some goals first.")
            return

        update_form = tk.Toplevel(self.root)
        update_form.title("Update Goal Progress")
        update_form.transient(self.root)
        update_form.grab_set()

        ttk.Label(update_form, text="Select Goal:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        goal_names = [f"{row['GoalID']} - {row['GoalName']}" for index, row in goals_df.iterrows()]
        goal_combo = ttk.Combobox(update_form, values=goal_names, state="readonly")
        goal_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        if goal_names:
            goal_combo.set(goal_names[0])

        ttk.Label(update_form, text="Amount to Add/Subtract:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        amount_entry = ttk.Entry(update_form)
        amount_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        amount_entry.insert(0, "0.00")

        def save_progress():
            selected_goal_str = goal_combo.get()
            if not selected_goal_str:
                messagebox.showwarning("Input Error", "Please select a goal.")
                return

            try:
                goal_id = int(selected_goal_str.split(' - ')[0])
            except ValueError:
                messagebox.showerror("Input Error", "Invalid goal selection.")
                return

            try:
                amount_change = float(amount_entry.get())
            except ValueError:
                messagebox.showerror("Input Error", "Please enter a valid number for amount.")
                return

            # Find the goal in the DataFrame
            idx_to_update = goals_df.index[goals_df['GoalID'] == goal_id].tolist()
            if not idx_to_update:
                messagebox.showerror("Error", "Selected goal not found.")
                return

            current_amount = goals_df.loc[idx_to_update[0], 'CurrentAmount']
            goals_df.loc[idx_to_update[0], 'CurrentAmount'] = max(0.0, current_amount + amount_change)

            # Update goal status if reached or exceeded
            if goals_df.loc[idx_to_update[0], 'CurrentAmount'] >= goals_df.loc[idx_to_update[0], 'TargetAmount']:
                goals_df.loc[idx_to_update[0], 'Status'] = 'Completed'
            elif goals_df.loc[idx_to_update[0], 'Status'] == 'Completed': # If it was completed but now below target
                goals_df.loc[idx_to_update[0], 'Status'] = 'In Progress'

            if update_table_data('Goals', goals_df):
                update_form.destroy()
                logging.info("Goal progress updated.")
                # Now calls the centralized refresh method
                self._refresh_all_tabs()
                messagebox.showinfo("Update Success", "Goal progress updated successfully. Data refreshed and synced to Excel.")
            else:
                messagebox.showerror("Update Failed", "Failed to update goal progress.")

        save_button = ttk.Button(update_form, text="Update Progress", command=save_progress)
        save_button.grid(row=2, column=0, columnspan=2, pady=10)

        update_form.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (update_form.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (update_form.winfo_height() // 2)
        update_form.geometry(f"+{x}+{y}")

    def on_closing(self):
        """Handles actions when the application window is closed."""
        logging.info("Application is closing. Performing final sync to Excel...")
        try:
            # Ensure all in-memory DataFrames are saved to DB before final sync
            update_account_balances_and_debt_amounts() # Final calculation update
            for table_name, df in self.data_frames.items():
                # Only save if the DataFrame exists and is not empty
                if df is not None and not df.empty:
                    update_table_data(table_name, df) # Save any unsaved changes from GUI operations
            sync_data('sqlite_to_excel') # Perform final sync from SQLite to Excel
            logging.info("Final sync completed. Destroying application window.")
            self.root.destroy()
        except Exception as e:
            logging.error(f"Error during final sync on close: {e}")
            messagebox.showerror("Exit Error", f"An error occurred during final data sync: {e}\nApplication will now close.")
            self.root.destroy() # Force close if sync fails

if __name__ == "__main__":
    # Ensure database is initialized before GUI attempts to load data
    logging.info("Ensuring database is initialized before GUI launch...")
    try:
        initialize_database()
        logging.info("Database initialization check complete.")
    except Exception as e:
        logging.error(f"Database initialization failed before GUI launch: {e}")
        messagebox.showerror("Startup Error", f"Failed to initialize database: {e}\nCannot launch application.")
        exit() # Exit if database cannot be initialized

    # Run the Tkinter application
    try:
        root = tk.Tk()
        app = DebtManagerApp(root)
        root.mainloop()
    except Exception as e:
        logging.critical(f"Unhandled exception during GUI application runtime: {e}", exc_info=True)
        messagebox.showerror("Application Error", f"An unexpected error occurred: {e}\nCheck DebugLog.txt for details.")

