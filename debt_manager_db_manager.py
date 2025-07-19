# debt_manager_db_manager.py
# Purpose: Manages all database interactions for the Debt Management System.
#          Handles SQLite database initialization, table creation, and CRUD operations.
# Deploy in: C:\DebtTracker
# Version: 1.4 (2025-07-19) - Updated get_table_data for Payments to include AccountID and AccountName.
#                            Adjusted update_account_balances to correctly use Payments.AccountID.
#                            Removed redundant internal initialize_database function.
#                            Now relies solely on debt_manager_db_init.py for schema management.

import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import os
import logging
from config import DB_PATH, TABLE_SCHEMAS, LOG_FILE, LOG_DIR, PREDEFINED_CATEGORIES

# Import the dedicated database initializer
from debt_manager_db_init import initialize_database as db_initializer_from_init_script

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

def get_db_connection():
    """Establishes and returns a connection to the SQLite database."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row # Allows accessing columns by name
        logging.debug("Database connection established.")
        return conn
    except sqlite3.Error as e:
        logging.critical(f"Database connection error: {e}", exc_info=True)
        raise

# Removed the redundant initialize_database function from here.
# Schema initialization is now handled exclusively by debt_manager_db_init.py.

def get_table_data(table_name):
    """
    Fetches all data from the specified table.
    Performs joins for specific tables to include display names for foreign keys.
    Returns data as a pandas DataFrame.
    """
    conn = None
    try:
        conn = get_db_connection()
        query = f"SELECT * FROM {table_name}"

        # Customize query for tables with foreign keys to fetch display names
        if table_name == 'Debts':
            query = """
                SELECT
                    d.DebtID, d.Creditor, d.OriginalAmount, d.Amount, d.AmountPaid,
                    d.MinimumPayment, d.SnowballPayment, d.InterestRate, d.DueDate,
                    d.Status, d.CategoryID, c.CategoryName, d.AccountID, a.AccountName
                FROM
                    Debts d
                LEFT JOIN
                    Categories c ON d.CategoryID = c.CategoryID
                LEFT JOIN
                    Accounts a ON d.AccountID = a.AccountID
            """
        elif table_name == 'Payments':
            # NEW: Include AccountID and AccountName for the payment source account
            query = """
                SELECT
                    p.PaymentID, p.DebtID, d.Creditor AS DebtName, p.AccountID, a.AccountName AS AccountName,
                    p.Amount, p.PaymentDate, p.PaymentMethod, p.CategoryID, c.CategoryName, p.Notes
                FROM
                    Payments p
                LEFT JOIN
                    Debts d ON p.DebtID = d.DebtID
                LEFT JOIN
                    Accounts a ON p.AccountID = a.AccountID -- NEW JOIN for payment source account
                LEFT JOIN
                    Categories c ON p.CategoryID = c.CategoryID
            """
        elif table_name == 'Revenue':
            query = """
                SELECT
                    r.RevenueID, r.Amount, r.DateReceived, r.Source,
                    r.AllocatedTo, r.AllocationType, r.NextProjectedIncome, r.NextProjectedIncomeDate,
                    r.AccountID, a.AccountName
                FROM
                    Revenue r
                LEFT JOIN
                    Accounts a ON r.AccountID = a.AccountID
            """
        elif table_name == 'Goals':
            query = """
                SELECT
                    g.GoalID, g.GoalName, g.TargetAmount, g.CurrentAmount, g.TargetDate,
                    g.Status, g.Notes, g.AccountID, a.AccountName
                FROM
                    Goals g
                LEFT JOIN
                    Accounts a ON g.AccountID = a.AccountID
            """
        elif table_name == 'Assets':
            query = """
                SELECT
                    a.AssetID, a.AssetName, a.Value, a.Category, a.PurchaseDate,
                    a.Status, a.Notes
                FROM
                    Assets a
            """ # No foreign keys to join for Assets based on schema

        df = pd.read_sql_query(query, conn)
        logging.debug(f"Loaded {len(df)} rows from {table_name}.")
        return df
    except sqlite3.Error as e:
        logging.error(f"Error loading data for {table_name}: {e}", exc_info=True)
        return pd.DataFrame() # Return empty DataFrame on error
    finally:
        if conn:
            conn.close()

def insert_record(table_name, data):
    """Inserts a new record into the specified table."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # Filter data to only include columns present in the schema for the table
        schema_columns = {col['name'] for col in TABLE_SCHEMAS[table_name]['columns']}
        insert_data = {k: v for k, v in data.items() if k in schema_columns}

        columns = ', '.join(insert_data.keys())
        placeholders = ', '.join(['?' for _ in insert_data.values()])
        sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

        cursor.execute(sql, tuple(insert_data.values()))
        conn.commit()
        logging.info(f"Added new record to {table_name}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Error adding record to {table_name}: {e}", exc_info=True)
        return False
    finally:
        if conn:
            conn.close()

def update_record(table_name, record_id, data):
    """Updates an existing record in the specified table."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        primary_key = TABLE_SCHEMAS[table_name]['primary_key']

        # Filter data to only include columns present in the schema for the table
        schema_columns = {col['name'] for col in TABLE_SCHEMAS[table_name]['columns']}
        update_data = {k: v for k, v in data.items() if k in schema_columns and k != primary_key}

        if not update_data:
            logging.warning(f"No valid data to update for {table_name} with ID={record_id}.")
            return False

        set_clauses = [f"{col} = ?" for col in update_data.keys()]
        sql = f"UPDATE {table_name} SET {', '.join(set_clauses)} WHERE {primary_key} = ?"

        values = list(update_data.values())
        values.append(record_id) # Add the ID for the WHERE clause

        logging.debug(f"Executing update for {table_name} ID {record_id}: SQL='{sql}', Values={values}")
        cursor.execute(sql, tuple(values))
        conn.commit()
        logging.info(f"Updated record in {table_name} with {primary_key}={record_id}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Error updating record in {table_name} with {primary_key}={record_id}: {e}", exc_info=True)
        return False
    finally:
        if conn:
            conn.close()

def delete_record(table_name, primary_key, record_id):
    """Deletes a record from the specified table."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        sql = f"DELETE FROM {table_name} WHERE {primary_key} = ?"
        cursor.execute(sql, (record_id,))
        conn.commit()
        logging.info(f"Deleted record from {table_name} with {primary_key}={record_id}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Error deleting record from {table_name} with {primary_key}={record_id}: {e}", exc_info=True)
        return False
    finally:
        if conn:
            conn.close()

def update_debt_amounts_and_payments():
    """
    Recalculates debt amounts and AmountPaid based on associated payments.
    This function is called after adding/editing/deleting payments.
    """
    logging.info("Updating debt amounts and AmountPaid based on payments.")
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # Get all debts
        debts_df = get_table_data('Debts')
        if debts_df.empty:
            logging.info("No debts to update debt amounts.")
            return

        # Get all payments
        payments_df = get_table_data('Payments')

        for index, debt_row in debts_df.iterrows():
            debt_id = debt_row['DebtID']
            original_amount = debt_row['OriginalAmount'] # Use OriginalAmount as base

            # Calculate total payments for this debt
            total_paid = payments_df[payments_df['DebtID'] == debt_id]['Amount'].sum()

            # Update AmountPaid and Amount (remaining balance)
            new_amount_paid = total_paid
            new_amount_remaining = original_amount - total_paid

            # Ensure Amount doesn't go below zero if payments exceed original amount
            if new_amount_remaining < 0:
                new_amount_remaining = 0

            # Update the debt record in the database
            update_sql = "UPDATE Debts SET AmountPaid = ?, Amount = ? WHERE DebtID = ?"
            cursor.execute(update_sql, (new_amount_paid, new_amount_remaining, debt_id))
            logging.debug(f"Updated DebtID {debt_id}: AmountPaid={new_amount_paid}, Amount={new_amount_remaining}")

        conn.commit()
        logging.info("Debt amounts and AmountPaid updated based on payments.")
    except sqlite3.Error as e:
        logging.error(f"Error updating debt amounts and payments: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

def update_account_balances():
    """
    Recalculates account balances based on associated revenue and payments.
    This function is called after adding/editing/deleting revenue or payments.
    """
    logging.info("Updating account balances.")
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        accounts_df = get_table_data('Accounts')
        if accounts_df.empty:
            logging.info("No accounts to update balances.")
            return

        revenue_df = get_table_data('Revenue')
        payments_df = get_table_data('Payments')

        for index, account_row in accounts_df.iterrows():
            account_id = account_row['AccountID']
            # Safely get 'InitialBalance' with a default if it doesn't exist yet
            initial_balance = account_row.get('InitialBalance', 0.0) # Use .get() for robustness

            # Sum revenue allocated directly to this account via AccountID
            total_revenue = revenue_df[revenue_df['AccountID'] == account_id]['Amount'].sum() if 'AccountID' in revenue_df.columns else 0

            # NEW: Sum payments made *from* this account (Payments.AccountID represents the source account)
            total_payments_from_account = payments_df[payments_df['AccountID'] == account_id]['Amount'].sum() if 'AccountID' in payments_df.columns else 0

            new_balance = initial_balance + total_revenue - total_payments_from_account

            update_sql = "UPDATE Accounts SET Balance = ? WHERE AccountID = ?"
            cursor.execute(update_sql, (new_balance, account_id))
            logging.debug(f"Updated AccountID {account_id}: Balance={new_balance}")

        conn.commit()
        logging.info("Account balances updated.")
    except sqlite3.Error as e:
        logging.error(f"Error updating account balances: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

def update_goal_progress():
    """
    Recalculates goal progress (CurrentAmount) based on associated revenue or payments.
    This function is called after adding/editing/deleting revenue or payments.
    """
    logging.info("Updating goal progress.")
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        goals_df = get_table_data('Goals')
        if goals_df.empty:
            logging.info("No goals to update progress.")
            return

        revenue_df = get_table_data('Revenue')
        # Payments could also contribute to goals, but current schema doesn't link directly.
        # Assuming Revenue is the primary source for updating goal CurrentAmount.

        for index, goal_row in goals_df.iterrows():
            goal_id = goal_row['GoalID']
            # Use AccountID if present in Goals table and Revenue can be allocated to it
            account_id = goal_row['AccountID'] if 'AccountID' in goal_row else None # Check if Goals has AccountID

            # Find total revenue allocated to this specific goal (via GoalName or AccountID)
            allocated_revenue = 0
            if 'AllocationType' in revenue_df.columns and 'AllocatedTo' in revenue_df.columns:
                # Prioritize direct allocation by GoalID/GoalName in Revenue.AllocatedTo
                allocated_revenue += revenue_df[
                    (revenue_df['AllocationType'] == 'Goal') &
                    (revenue_df['AllocatedTo'] == str(goal_row['GoalName'])) # Compare with string representation of GoalName
                ]['Amount'].sum()

                # Also consider revenue directly linked to the account associated with the goal
                if account_id is not None and 'AccountID' in revenue_df.columns:
                    allocated_revenue += revenue_df[
                        (revenue_df['AllocationType'] == 'Account') & # Assuming revenue allocated to account type
                        (revenue_df['AccountID'] == account_id)
                    ]['Amount'].sum()

            # For now, let's just sum revenue directly linked to AccountID, if present.
            # This logic needs to align with how you intend revenue to populate goals.
            # If Revenue has AccountID, and Goal has AccountID, we can link them.
            if 'AccountID' in revenue_df.columns and account_id is not None:
                 # Sum revenue that is associated with the account linked to this goal
                 # This assumes that revenue meant for a goal is recorded as income to the associated account
                allocated_revenue = revenue_df[revenue_df['AccountID'] == account_id]['Amount'].sum()
            else:
                allocated_revenue = 0 # If no direct account link or revenue does not have AccountID

            new_current_amount = allocated_revenue # Or goal_row['CurrentAmount'] + allocated_revenue if it's incremental

            update_sql = "UPDATE Goals SET CurrentAmount = ? WHERE GoalID = ?"
            cursor.execute(update_sql, (new_current_amount, goal_id))
            logging.debug(f"Updated GoalID {goal_id}: CurrentAmount={new_current_amount}")

        conn.commit()
        logging.info("Goal progress updated.")
    except sqlite3.Error as e:
        logging.error(f"Error updating goal progress: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

def generate_financial_projection(start_date=None, num_years=20):
    """
    Generates a financial projection based on current debts, payments, and revenue.
    This is a simplified model for demonstration.
    """
    logging.info("Generating financial projection.")

    # Ensure database is initialized before fetching data
    # The orchestrator calls db_init.py explicitly, so this is for standalone calls or robustness.
    db_initializer_from_init_script() # Call the initializer from db_init.py


    debts_df = get_table_data('Debts')
    accounts_df = get_table_data('Accounts')
    revenue_df = get_table_data('Revenue')

    if debts_df.empty and accounts_df.empty and revenue_df.empty:
        logging.warning("No financial data to generate projection.")
        return pd.DataFrame()

    current_date = datetime.now()
    if start_date:
        try:
            current_date = datetime.strptime(str(start_date), '%Y-%m-%d')
        except ValueError:
            logging.error(f"Invalid start_date format: {start_date}. Using current date.")
            current_date = datetime.now()

    projection_data = []

    # Initial state
    total_debt = debts_df[debts_df['Status'] != 'Paid Off']['Amount'].sum() if not debts_df.empty else 0
    total_savings = accounts_df[accounts_df['AccountType'].isin(['Checking', 'Savings', 'Investment'])]['Balance'].sum() if not accounts_df.empty else 0
    net_worth = total_savings - total_debt

    projection_data.append({
        'Year': current_date.year,
        'DebtRemaining': total_debt,
        'Savings': total_savings,
        'NetWorth': net_worth
    })

    # Simulate month by month
    sim_debt = total_debt
    sim_savings = total_savings

    # Get total minimum payments and snowball payments
    total_min_payments = debts_df[debts_df['Status'] != 'Paid Off']['MinimumPayment'].sum() if not debts_df.empty else 0
    total_snowball_payments = debts_df[debts_df['Status'] != 'Paid Off']['SnowballPayment'].sum() if not debts_df.empty else 0

    # Assuming average monthly revenue
    # Check if 'NextProjectedIncome' and 'NextProjectedIncomeDate' exist before using
    avg_monthly_projected_revenue = revenue_df['NextProjectedIncome'].sum() if 'NextProjectedIncome' in revenue_df.columns and not revenue_df.empty else 0

    # Or, if you want average of historical revenue:
    # avg_monthly_revenue = revenue_df['Amount'].mean() / 12 if not revenue_df.empty else 0

    for year_offset in range(1, num_years + 1):
        for month_offset in range(1, 13):
            # Simulate monthly changes
            # Revenue inflow
            sim_savings += avg_monthly_projected_revenue # Using projected monthly income

            # Debt payments (prioritize minimum, then snowball)
            if sim_debt > 0:
                payment_this_month = total_min_payments + total_snowball_payments

                # Ensure we don't pay more than remaining debt
                if payment_this_month > sim_debt:
                    payment_this_month = sim_debt

                sim_debt -= payment_this_month
                sim_savings -= payment_this_month # Money leaves savings to pay debt

                if sim_debt < 0:
                    sim_debt = 0 # Debt cannot be negative

            # Interest accumulation (simplified)
            # For a real projection, this would be per-debt and more complex
            if sim_debt > 0:
                sim_debt *= (1 + (0.05 / 12)) # Assume 5% average annual interest

            # Savings growth (simplified)
            sim_savings *= (1 + (0.01 / 12)) # Assume 1% average annual savings interest

        # After each year, record the state
        projection_data.append({
            'Year': current_date.year + year_offset,
            'DebtRemaining': sim_debt,
            'Savings': sim_savings,
            'NetWorth': sim_savings - sim_debt
        })

    projection_df = pd.DataFrame(projection_data)
    logging.info(f"Financial projection generated with {len(projection_df)} rows.")
    return projection_df

# Schema update function (can be called separately if needed)
# This function is now considered a helper for cases where only updates are needed
# without full DB initialization/recreation. The primary schema management is in db_init.py.
def update_database_schema():
    """
    Updates the database schema by adding any missing columns defined in TABLE_SCHEMAS.
    This is useful for applying schema changes without recreating the entire database.
    """
    logging.info("Starting database schema update process.")
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        for table_name, schema in TABLE_SCHEMAS.items():
            logging.info(f"Table '{table_name}' already exists. Checking for missing columns.")
            cursor.execute(f"PRAGMA table_info({table_name})")
            existing_columns = [col[1] for col in cursor.fetchall()]

            for col in schema['columns']:
                if col['name'] not in existing_columns:
                    add_column_sql = f"ALTER TABLE {table_name} ADD COLUMN {col['name']} {col['type']}"
                    if 'default' in col:
                        default_val = col['default']
                        if isinstance(default_val, str):
                            add_column_sql += f" DEFAULT '{default_val}'"
                        else:
                            add_column_sql += f" DEFAULT {default_val}"
                    try:
                        cursor.execute(add_column_sql)
                        conn.commit()
                        logging.info(f"Added missing column to {table_name}: {col['name']} {col['type']}.")
                    except sqlite3.Error as e:
                        logging.error(f"Error adding column {col['name']} to {table_name}: {e}")
        conn.commit()
        logging.info("Database schema update process completed.")
    except sqlite3.Error as e:
        logging.critical(f"An error occurred during database schema update: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    # This block is for testing the db_manager functions directly
    # It will not run when imported by other scripts like the GUI.

    # Initialize database (creates if not exists, adds tables, adds categories)
    # Use the dedicated db_init script's initialize_database for comprehensive setup
    db_initializer_from_init_script()

    # Example: Add sample data if tables are empty
    # Check if 'InitialBalance' column exists in Accounts before trying to use it for sample data
    conn_check = get_db_connection()
    cursor_check = conn_check.cursor()
    cursor_check.execute("PRAGMA table_info(Accounts)")
    accounts_cols = [col[1] for col in cursor_check.fetchall()]
    conn_check.close()

    if get_table_data('Debts').empty:
        logging.info("Adding sample debt.")
        insert_record('Debts', {
            'Creditor': 'Sample Credit Card',
            'OriginalAmount': 1000.0,
            'Amount': 1000.0,
            'AmountPaid': 0.0,
            'MinimumPayment': 25.0,
            'SnowballPayment': 0.0,
            'InterestRate': 18.0,
            'DueDate': '2025-12-31',
            'Status': 'Open',
            'CategoryID': 1, # Assuming 'Credit Card' category exists
            'AccountID': None # No account linked initially
        })

    if get_table_data('Accounts').empty:
        logging.info("Adding sample account.")
        sample_account_data = {
            'AccountName': 'Main Checking',
            'Balance': 1500.0,
            'AccountType': 'Checking',
            'Status': 'Open',
        }
        if 'InitialBalance' in accounts_cols: # Only add if column exists
            sample_account_data['InitialBalance'] = 1500.0
        insert_record('Accounts', sample_account_data)

    # Example: Add a sample payment for the sample debt
    if get_table_data('Payments').empty and not get_table_data('Debts').empty:
        logging.info("Adding sample payment.")
        debt_id = get_table_data('Debts')['DebtID'].iloc[0] # Get the first debt ID
        account_id = get_table_data('Accounts')['AccountID'].iloc[0] # Get the first account ID
        # Assuming CategoryID 7 for 'Debt Payment' based on PREDEFINED_CATEGORIES order
        insert_record('Payments', {
            'DebtID': debt_id,
            'AccountID': account_id, # NEW: Link to a sample account
            'Amount': 50.0,
            'PaymentDate': '2025-07-15',
            'PaymentMethod': 'Bank Transfer',
            'CategoryID': 7, # Assuming 'Debt Payment' is CategoryID 7 based on PREDEFINED_CATEGORIES order
            'Notes': 'Initial payment',
        })

    # Example: Add a sample revenue
    if get_table_data('Revenue').empty:
        logging.info("Adding sample revenue.")
        # Assuming AccountID 1 for 'Main Checking'
        insert_record('Revenue', {
            'Amount': 1000.0,
            'DateReceived': '2025-07-10',
            'Source': 'Salary',
            'AllocatedTo': 1, # Assuming AccountID 1 for 'Main Checking'
            'AllocationType': 'Account',
            'NextProjectedIncome': 1000.0,
            'NextProjectedIncomeDate': '2025-08-10',
            'AccountID': 1 # Direct link to Account for convenience
        })

    # Example: Add a sample goal
    if get_table_data('Goals').empty:
        logging.info("Adding sample goal.")
        insert_record('Goals', {
            'GoalName': 'Emergency Fund',
            'TargetAmount': 5000.0,
            'CurrentAmount': 0.0,
            'TargetDate': '2026-12-31',
            'Status': 'Planned',
            'Notes': 'Build up savings for emergencies.',
            'AccountID': 1 # Link to a sample account, e.g., 'Main Checking'
        })


    # Test update functions
    update_debt_amounts_and_payments()
    update_account_balances()
    update_goal_progress()

    logging.info("DB Manager test run completed.")
