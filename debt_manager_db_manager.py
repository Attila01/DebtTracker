# debt_manager_db_manager.py
# Purpose: Manages all interactions with the SQLite database (debt_manager.db).
#          Provides functions for CRUD operations, balance updates, and projections.
# Deploy in: C:\DebtTracker
# Version: 1.0 (2025-07-19) - Initial version. Python-based database manager.

import sqlite3
import os
import logging
import pandas as pd
from datetime import datetime, timedelta

from config import DB_PATH, TABLE_SCHEMAS, LOG_FILE, LOG_DIR, PREDEFINED_CATEGORIES

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
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row # Allows accessing columns by name
        logging.debug("SQLite database connection opened.")
        return conn
    except sqlite3.Error as e:
        logging.error(f"Error connecting to database: {e}", exc_info=True)
        return None

def initialize_database():
    """
    Initializes the SQLite database: creates the database file if it doesn't exist,
    and ensures all necessary tables are created with the correct schema.
    Also inserts predefined categories if the Categories table is empty.
    """
    logging.info("Starting database initialization process.")

    # Ensure database directory exists
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

    conn = None
    try:
        conn = get_db_connection()
        if not conn:
            raise Exception("Failed to get database connection.")
        cursor = conn.cursor()

        # Check if the database file is a valid SQLite database
        # This is a basic check; a more robust check might involve PRAGMA integrity_check
        if os.path.exists(DB_PATH) and os.path.getsize(DB_PATH) > 0:
            try:
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
                logging.info(f"Existing database file '{DB_PATH}' is a valid SQLite database.")
            except sqlite3.DatabaseError as e:
                logging.critical(f"Critical database initialization error: {e}")
                logging.critical(f"The file at '{DB_PATH}' exists but is not a valid SQLite database. "
                                 "It might be an Access database (.accdb) or corrupted. "
                                 "Please delete it or rename it if you want to create a new SQLite database.")
                raise Exception(f"Invalid database file: {DB_PATH}. {e}")

        for table_name, schema in TABLE_SCHEMAS.items():
            columns_sql = []
            primary_key_cols = []
            for col in schema['columns']:
                col_def = f"{col['name']} {col['type']}"
                if col.get('primary_key'):
                    primary_key_cols.append(col['name'])
                    if col.get('autoincrement'):
                        col_def += ' PRIMARY KEY AUTOINCREMENT' # SQLite specific auto-increment
                if not col.get('nullable') and not col.get('primary_key'):
                    col_def += ' NOT NULL'
                if 'default' in col:
                    default_val = col['default']
                    if isinstance(default_val, str):
                        col_def += f" DEFAULT '{default_val}'"
                    else:
                        col_def += f" DEFAULT {default_val}"
                if col.get('unique'):
                    col_def += ' UNIQUE'
                columns_sql.append(col_def)

            # Check if table exists
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
            if cursor.fetchone():
                logging.info(f"Table '{table_name}' already exists. Checking for missing columns.")
                # Update schema by adding missing columns
                cursor.execute(f"PRAGMA table_info({table_name});")
                existing_columns = [info[1] for info in cursor.fetchall()] # info[1] is column name

                for col in schema['columns']:
                    if col['name'] not in existing_columns:
                        col_def = f"{col['name']} {col['type']}"
                        if not col.get('nullable') and not col.get('primary_key'):
                            col_def += ' NOT NULL'
                        if 'default' in col:
                            default_val = col['default']
                            if isinstance(default_val, str):
                                col_def += f" DEFAULT '{default_val}'"
                            else:
                                col_def += f" DEFAULT {default_val}"
                        if col.get('unique'):
                            col_def += ' UNIQUE'

                        try:
                            cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_def};")
                            logging.info(f"Adding missing column to {table_name}: {col_def}")
                            conn.commit()
                            logging.info(f"Column '{col['name']}' added to table '{table_name}'.")
                        except sqlite3.Error as e:
                            logging.warning(f"Could not add column {col['name']} to {table_name}: {e}")
            else:
                create_table_sql = f"CREATE TABLE {table_name} ({', '.join(columns_sql)});"
                cursor.execute(create_table_sql)
                conn.commit()
                logging.info(f"Table '{table_name}' created successfully.")

        # Special handling for Categories table: insert predefined categories if empty
        if table_name == 'Categories':
            cursor.execute("SELECT COUNT(*) FROM Categories")
            if cursor.fetchone()[0] == 0:
                logging.info("Categories table is empty. Inserting predefined categories.")
                for category_name in PREDEFINED_CATEGORIES:
                    try:
                        cursor.execute("INSERT INTO Categories (CategoryName) VALUES (?)", (category_name,))
                        conn.commit()
                    except sqlite3.IntegrityError: # In case of unique constraint violation (shouldn't happen if table was empty)
                        logging.warning(f"Category '{category_name}' already exists, skipping insertion.")
                logging.info("Predefined categories inserted successfully.")
            else:
                logging.info("Categories table already contains data. Skipping predefined category insertion.")

        logging.info("Database initialization process completed.")

    except sqlite3.Error as e:
        logging.critical(f"Database initialization failed: {e}", exc_info=True)
        raise
    except Exception as e:
        logging.critical(f"An unexpected error occurred during database initialization: {e}", exc_info=True)
        raise
    finally:
        if conn:
            conn.close()
            logging.info("SQLite connection closed after initialization.")

def get_table_data(table_name):
    """Fetches all data from a specified table."""
    conn = get_db_connection()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        logging.debug(f"Loaded {table_name} ({len(df)} rows)")
        return df
    except sqlite3.Error as e:
        logging.error(f"Error loading data from {table_name}: {e}", exc_info=True)
        return pd.DataFrame()
    finally:
        if conn:
            conn.close()

def insert_record(table_name, data):
    """Inserts a new record into the specified table."""
    conn = get_db_connection()
    if not conn:
        return False
    try:
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?' for _ in data.values()])
        sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
        cursor = conn.cursor()
        cursor.execute(sql, tuple(data.values()))
        conn.commit()
        logging.info(f"Added new record to {table_name}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Error adding record to {table_name}: {e}", exc_info=True)
        return False
    finally:
        if conn:
            conn.close()

def update_record(table_name, primary_key_value, data):
    """Updates an existing record in the specified table."""
    conn = get_db_connection()
    if not conn:
        return False
    try:
        schema = TABLE_SCHEMAS.get(table_name)
        if not schema:
            logging.error(f"Schema not found for table: {table_name}")
            return False

        primary_key_col = schema['primary_key']

        set_clauses = []
        values = []
        for col, val in data.items():
            if col != primary_key_col: # Don't update primary key
                set_clauses.append(f"{col} = ?")
                values.append(val)

        values.append(primary_key_value) # Add primary key value for WHERE clause

        sql = f"UPDATE {table_name} SET {', '.join(set_clauses)} WHERE {primary_key_col} = ?"
        cursor = conn.cursor()
        cursor.execute(sql, tuple(values))
        conn.commit()
        logging.info(f"Updated record in {table_name} with {primary_key_col}={primary_key_value}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Error updating record in {table_name} with {primary_key_col}={primary_key_value}: {e}", exc_info=True)
        return False
    finally:
        if conn:
            conn.close()

def delete_record(table_name, primary_key_col, primary_key_value):
    """Deletes a record from the specified table."""
    conn = get_db_connection()
    if not conn:
        return False
    try:
        sql = f"DELETE FROM {table_name} WHERE {primary_key_col} = ?"
        cursor = conn.cursor()
        cursor.execute(sql, (primary_key_value,))
        conn.commit()
        logging.info(f"Deleted record from {table_name} with {primary_key_col}={primary_key_value}.")
        return True
    except sqlite3.Error as e:
        logging.error(f"Error deleting record from {table_name} with {primary_key_col}={primary_key_value}: {e}", exc_info=True)
        return False
    finally:
        if conn:
            conn.close()

def update_debt_amounts_and_payments():
    """
    Updates 'Amount' and 'AmountPaid' in Debts table based on Payments.
    Also updates Account balances.
    """
    conn = get_db_connection()
    if not conn:
        return

    try:
        cursor = conn.cursor()

        # Update AmountPaid for each debt
        cursor.execute("""
            UPDATE Debts
            SET AmountPaid = (
                SELECT SUM(P.Amount)
                FROM Payments P
                WHERE P.DebtID = Debts.DebtID
            )
            WHERE EXISTS (SELECT 1 FROM Payments WHERE Payments.DebtID = Debts.DebtID)
        """)
        # Set AmountPaid to 0 if no payments exist for a debt
        cursor.execute("""
            UPDATE Debts
            SET AmountPaid = 0.0
            WHERE NOT EXISTS (SELECT 1 FROM Payments WHERE Payments.DebtID = Debts.DebtID)
        """)
        conn.commit()
        logging.info("Debt amounts and AmountPaid updated based on payments.")

        # Update 'Amount' (remaining debt)
        cursor.execute("""
            UPDATE Debts
            SET Amount = OriginalAmount - AmountPaid
            WHERE OriginalAmount IS NOT NULL AND AmountPaid IS NOT NULL
        """)
        conn.commit()
        logging.info("Remaining debt amounts updated.")

        # Now update account balances
        update_account_balances()

    except sqlite3.Error as e:
        logging.error(f"Error updating debt amounts and payments: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

def update_account_balances():
    """Recalculates and updates the Balance for all accounts."""
    conn = get_db_connection()
    if not conn:
        return

    try:
        cursor = conn.cursor()

        accounts_df = get_table_data('Accounts')
        if accounts_df.empty:
            logging.info("No accounts to update balances.")
            return

        for index, account in accounts_df.iterrows():
            account_id = account['AccountID']

            # Calculate deposits from Revenue allocated to this account
            cursor.execute("""
                SELECT SUM(Amount) FROM Revenue
                WHERE AllocatedTo = ? AND AllocationType = 'Account'
            """, (account_id,))
            deposits = cursor.fetchone()[0] or 0.0

            # Calculate withdrawals from Payments linked to this account (as a 'debt' for simplicity)
            # This assumes that if a payment is linked to an AccountID, it's a withdrawal from that account.
            # This might need refinement based on how you categorize payments (e.g., bills vs transfers)
            cursor.execute("""
                SELECT SUM(Amount) FROM Payments
                WHERE DebtID = ? -- Assuming DebtID can also refer to AccountID for withdrawals
            """, (account_id,))
            withdrawals = cursor.fetchone()[0] or 0.0

            # For credit card accounts, payments *to* the credit card are deposits to the account balance,
            # and charges *on* the credit card are withdrawals.
            # This logic needs to be more sophisticated if 'DebtID' in Payments is strictly for debts.
            # For now, let's assume 'DebtID' in Payments can also refer to AccountID for simplicity.
            # A better approach would be:
            # - Payments *from* this account (e.g., for debts)
            # - Revenue *to* this account (deposits)
            # - Direct charges/credits to this account (requires a 'Transactions' table)

            # For now, let's simplify:
            # Balance = Initial Balance + Deposits (Revenue to Account) - Payments (from Account)
            # If AccountType is 'Credit Card', then it's more complex:
            # Current Balance = Starting Limit - Payments + Charges

            # Re-evaluating based on original PowerShell logic:
            # $deposits = SUM(Revenue.Amount WHERE AllocatedTo = AccountID AND AllocationType = 'Account')
            # $withdrawals = SUM(Payments.Amount WHERE DebtID = AccountID) -- This implies payments *from* the account
            # This seems to treat accounts as "debts" in the Payments table, which is unusual.
            # Let's stick to a more conventional financial model for Python:
            # Account Balance = Initial Balance + Revenue (deposits) - Payments (withdrawals)
            # For credit cards, it's usually: Current Balance = Previous Balance + New Charges - Payments

            # Let's use a simpler interpretation:
            # Balance = starting_balance (if available) + total_revenue_allocated_to_account - total_payments_from_account
            # Since we don't have a 'starting_balance' in the schema, we'll calculate based on transactions.
            # This implies 'Balance' in the Accounts table is a calculated field.

            # For now, let's use the PowerShell logic's interpretation:
            # Deposits are Revenue where AllocationType is 'Account' and AllocatedTo is this AccountID.
            # Withdrawals are Payments where DebtID is this AccountID (meaning payment *from* this account for something else).
            # This is a bit of a hacky interpretation if DebtID is strictly for debts.
            # A proper transaction system would be better.

            # Let's use the direct sum of payments *from* this account (if we assume Payment.DebtID can be AccountID)
            # and sum of revenue *to* this account.

            # Sum of all revenue amounts where this account is the target of allocation
            cursor.execute("""
                SELECT SUM(Amount) FROM Revenue
                WHERE AllocatedTo = ? AND AllocationType = 'Account'
            """, (account_id,))
            total_allocated_revenue = cursor.fetchone()[0] or 0.0

            # Sum of all payments where this account is the 'source' (if DebtID can represent source account)
            # This interpretation is problematic. A 'SourceAccountID' in Payments would be better.
            # Sticking to the original PowerShell interpretation of Payment.DebtID referring to AccountID for withdrawals.
            cursor.execute("""
                SELECT SUM(Amount) FROM Payments
                WHERE DebtID = ?
            """, (account_id,))
            total_payments_from_account = cursor.fetchone()[0] or 0.0

            # The 'Balance' field in Accounts is the running balance.
            # This calculation is simplistic and might not perfectly model all account types.
            # For a credit card, a "payment" to it increases its balance (reduces debt).
            # For a checking account, a "payment" from it decreases its balance.
            # The current schema doesn't clearly distinguish this.
            # For now, we'll assume 'Balance' is a net value.

            # If AccountType is 'Credit Card', the "Balance" is usually how much is owed.
            # Payments to a credit card REDUCE the amount owed.
            # If AccountType is 'Checking'/'Savings', payments REDUCE the balance.

            # Let's refine based on typical financial app logic:
            # For Checking/Savings: Balance = Initial + Deposits - Withdrawals
            # For Credit Card: Balance = Initial + Charges - Payments
            # Since we don't have a transaction table, we'll use the existing fields.

            # Simplified logic based on existing schema and common sense for "Balance":
            # If it's a Credit Card, the 'Balance' field likely means 'amount owed'.
            # Payments to this 'DebtID' (which is the Credit Card AccountID) should reduce this 'Balance'.
            # Revenue allocated to this 'AccountID' (e.g., a refund) should also reduce the 'Balance'.

            # Let's assume 'Balance' in Accounts is the *current* balance.
            # For non-credit accounts: Balance = (previous balance) + revenue - payments
            # For credit accounts: Balance = (previous balance) + charges - payments

            # Given the lack of a transaction log, we'll calculate based on sums:
            # Net effect on account = Sum of Revenue where AllocatedTo = AccountID
            #                       - Sum of Payments where DebtID = AccountID (assuming this means payments *from* account)
            # This is still problematic as Payments.DebtID is designed for actual debts.

            # Reverting to the PowerShell script's implicit logic for now, as it's what the user had:
            # $deposits = SUM(Revenue.Amount WHERE AllocatedTo = AccountID AND AllocationType = 'Account')
            # $withdrawals = SUM(Payments.Amount WHERE DebtID = AccountID)
            # New Balance = deposits - withdrawals (This implies accounts start at 0 or are reset)
            # This is only valid if 'Balance' is always calculated from scratch.
            # The PowerShell script updates existing balance, implying it's a running total.

            # Let's try to infer a running balance:
            # Get current balance from DB first
            cursor.execute("SELECT Balance FROM Accounts WHERE AccountID = ?", (account_id,))
            current_balance_in_db = cursor.fetchone()[0] or 0.0

            # Sum of all relevant transactions
            # Sum of revenue allocated to this account
            cursor.execute("""
                SELECT SUM(Amount) FROM Revenue
                WHERE AllocatedTo = ? AND AllocationType = 'Account'
            """, (account_id,))
            revenue_for_account = cursor.fetchone()[0] or 0.0

            # Sum of payments *from* this account (if we treat DebtID as source account for simplicity, or if it's a payment to a credit account)
            # This is still a weak link. The original PowerShell `Update-AccountBalances` uses `Payments WHERE DebtID = AccountID`.
            # This is ambiguous. Is a payment *to* a credit card a "payment" where DebtID is the credit card's AccountID?
            # Or is it a payment *from* a checking account where DebtID is the checking account's AccountID?
            # Let's assume for now that if Payments.DebtID matches an AccountID, it's a transaction *affecting* that account.
            # If AccountType is 'Credit Card', payments REDUCE the amount owed (increase "balance" in a positive sense).
            # If AccountType is 'Checking'/'Savings', payments REDUCE the balance.

            # Given the ambiguity, the simplest interpretation of the PowerShell code's `deposits - withdrawals`
            # is that it calculates the *net change* and sets the balance to that.
            # This means 'Balance' is not a running total, but a sum of transactions.

            # Let's stick to the direct translation of the PowerShell logic for now:
            # Deposits are revenue allocated to the account.
            # Withdrawals are payments where the DebtID matches the AccountID.
            # New Balance = Deposits - Withdrawals (this is a *net position*, not a running balance from an initial state)

            # This is the direct translation of the PowerShell logic:
            cursor.execute("""
                SELECT SUM(Amount) FROM Revenue WHERE AllocatedTo = ? AND AllocationType = 'Account'
            """, (account_id,))
            deposits = cursor.fetchone()[0] or 0.0

            cursor.execute("""
                SELECT SUM(Amount) FROM Payments WHERE DebtID = ?
            """, (account_id,))
            withdrawals = cursor.fetchone()[0] or 0.0

            new_balance = deposits - withdrawals

            cursor.execute("UPDATE Accounts SET Balance = ? WHERE AccountID = ?", (new_balance, account_id))
            conn.commit()
            logging.info(f"Updated balance for AccountID {account_id} to {new_balance}.")

        logging.info("Account balances updated.")

    except sqlite3.Error as e:
        logging.error(f"Error updating account balances: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

def update_goal_progress():
    """Recalculates and updates the CurrentAmount and Status for all goals."""
    conn = get_db_connection()
    if not conn:
        return

    try:
        cursor = conn.cursor()

        goals_df = get_table_data('Goals')
        if goals_df.empty:
            logging.info("No goals to update progress.")
            return

        for index, goal in goals_df.iterrows():
            goal_id = goal['GoalID']
            target_amount = goal['TargetAmount']

            # Sum payments categorized as 'Debt Payment' up to today
            # This logic seems to tie goal progress to debt payments, which might not be universal for all goals.
            # The original PowerShell was: SUM(Payments.Amount WHERE Category = 'Debt Payment' AND PaymentDate <= Get-Date)
            # This implies all goals are funded by "debt payments" which is odd.
            # A more flexible approach would be to link goals to specific revenue or asset increases.
            # For now, adhering to the original logic:
            cursor.execute("""
                SELECT SUM(Amount) FROM Payments
                WHERE Category = 'Debt Payment' AND PaymentDate <= ?
            """, (datetime.now().strftime('%Y-%m-%d'),))
            progress = cursor.fetchone()[0] or 0.0

            status = 'In Progress'
            if progress >= target_amount:
                status = 'Completed'
            elif progress == 0:
                status = 'Planned' # Or 'Not Started'

            cursor.execute("""
                UPDATE Goals SET CurrentAmount = ?, Status = ? WHERE GoalID = ?
            """, (progress, status, goal_id))
            conn.commit()
            logging.info(f"Updated progress for GoalID {goal_id} to {progress} ({status}).")

        logging.info("Goal progress updated.")

    except sqlite3.Error as e:
        logging.error(f"Error updating goal progress: {e}", exc_info=True)
    finally:
        if conn:
            conn.close()

def generate_financial_projection():
    """Generates a financial projection report based on debts, savings, and income."""
    conn = get_db_connection()
    if not conn:
        return pd.DataFrame()

    try:
        cursor = conn.cursor()

        # Total Debt
        cursor.execute("SELECT SUM(Amount) FROM Debts WHERE Status NOT IN ('Paid Off', 'Closed')")
        total_debt = cursor.fetchone()[0] or 0.0

        # Total Savings (from Accounts with appropriate type)
        cursor.execute("SELECT SUM(Balance) FROM Accounts WHERE AccountType IN ('Checking', 'Savings', 'Investment') AND Status IN ('Open', 'Current', 'Active')")
        total_savings = cursor.fetchone()[0] or 0.0

        # Annual Income (from Revenue over last 12 months)
        one_year_ago = (datetime.now() - timedelta(days=365)).strftime('%Y-%m-%d')
        cursor.execute("SELECT SUM(Amount) FROM Revenue WHERE DateReceived >= ?", (one_year_ago,))
        annual_income = cursor.fetchone()[0] or 0.0

        # Debts for snowball calculation (ordered by amount ascending)
        debts_df = get_table_data('Debts')
        debts_for_snowball = debts_df[
            (debts_df['Status'] != 'Paid Off') &
            (debts_df['Status'] != 'Closed')
        ].sort_values(by='Amount', ascending=True)

        years = [3, 5, 7, 10]
        projections = []

        for year in years:
            months = year * 12
            remaining_debt_at_projection = total_debt
            snowball_payment_pool = 0.0 # This will accumulate minimum payments from paid-off debts

            # Simulate debt repayment
            current_debts = debts_for_snowball.copy() # Work on a copy

            for index, debt in current_debts.iterrows():
                debt_amount = debt['Amount']
                minimum_payment = debt['MinimumPayment'] or 0.0
                snowball_contribution = debt['SnowballPayment'] or 0.0 # This is the user-defined snowball for this debt

                # Calculate effective monthly payment for this debt
                # This logic assumes the snowball payment is *additional* to minimums
                # and is applied to the smallest debt first.

                # The original PowerShell logic for snowball was:
                # $monthlyPayment = $minimumPayment + $snowball
                # $debtPaid = [Math]::Min($debtAmount, $monthlyPayment * $months)
                # if ($debtPaid -ge $debtAmount) { $snowball += $minimumPayment }
                # This means '$snowball' in PS was the *accumulated* snowball from previous debts.
                # Let's adjust for this interpretation.

                # Recalculate remaining debt and snowball based on the PS logic
                # This requires a more iterative monthly simulation or a more complex formula.
                # For simplicity and direct translation of the PS logic, which is a bit abstract:
                # The PS script's snowball logic is simplified and applies the 'snowball'
                # as a general pool that grows.

                # Let's re-implement the PS snowball logic more carefully:
                # The PS script iterates through debts ordered by amount.
                # When a debt is paid off, its minimum payment is added to the 'snowball' pool.
                # This 'snowball' pool is then added to the monthly payment of the *next* debt.

                # Simplified simulation for projection:
                # This is a high-level projection, not a detailed month-by-month simulation.
                # It estimates how much debt could be paid off in 'months' given minimums + a general snowball strategy.

                # Let's use a simplified model for the projection, similar to the PS script:
                # Assume a fixed monthly payment capacity that includes minimums + an extra snowball amount.
                # The PS script's snowball logic is somewhat abstract for a projection over multiple years without
                # a detailed monthly simulation.

                # Let's take the approach of the PS script's `New-FinancialProjection` directly:
                # It iterates through debts, and if a debt is "paid" within the projection period,
                # its minimum payment is added to a `snowball` variable for the *next* debt's calculation.

                simulated_remaining_debt = total_debt # Start with total debt
                current_snowball_pool = 0.0 # Accumulated minimum payments from paid-off debts

                # Sort debts by Amount ASC for snowball method
                sorted_debts = debts_df[
                    (debts_df['Status'] != 'Paid Off') &
                    (debts_df['Status'] != 'Closed')
                ].sort_values(by='Amount', ascending=True).to_dict(orient='records')

                # Calculate total minimum payments
                total_minimum_payments = sum(d['MinimumPayment'] for d in sorted_debts if d['MinimumPayment'] is not None)

                # This is a very simplified projection. It assumes all minimum payments are made
                # and then a snowball amount is applied.
                # The PS script's snowball logic is quite simplified for a multi-year projection.
                # It basically calculates if a debt *could* be paid off within the time frame
                # by simply multiplying monthly payment by months.

                # Let's re-align with the PS `New-FinancialProjection` logic:
                # It calculates a `remainingDebt` by simulating payment for each debt.
                # It uses `totalDebt` as the starting point.
                # The `$snowball` variable in PS is the *extra* amount applied to the current debt
                # from previous paid-off debts' minimum payments.

                temp_total_debt = total_debt
                temp_snowball_pool = 0.0

                for debt in sorted_debts:
                    debt_amount = debt['Amount'] or 0.0
                    min_payment = debt['MinimumPayment'] or 0.0

                    # Monthly payment for this specific debt, including accumulated snowball from previous debts
                    monthly_payment_for_this_debt = min_payment + temp_snowball_pool

                    # Amount paid on this debt over the projection period
                    # This is a very rough estimate as it doesn't account for interest or actual monthly payments
                    # but rather total capacity over the period.
                    amount_paid_on_this_debt = min(debt_amount, monthly_payment_for_this_debt * months)

                    temp_total_debt -= amount_paid_on_this_debt

                    # If this debt is 'paid off' within the projection, add its minimum payment to the snowball pool
                    if amount_paid_on_this_debt >= debt_amount and min_payment > 0:
                        temp_snowball_pool += min_payment

                remaining_debt_after_projection = max(0, temp_total_debt) # Ensure it doesn't go negative

                # Savings projection (PS: $totalSavings * [Math]::Pow(1 + 0.05, $year) + ($annualIncome * 0.2 * $year))
                # Assuming 5% annual growth on savings and 20% of annual income added to savings.
                projected_savings = total_savings * (1 + 0.05)**year + (annual_income * 0.2 * year)

                net_worth = projected_savings - remaining_debt_after_projection

                projections.append({
                    'Year': year,
                    'DebtRemaining': round(remaining_debt_after_projection, 2),
                    'Savings': round(projected_savings, 2),
                    'NetWorth': round(net_worth, 2)
                })

        projections_df = pd.DataFrame(projections)
        logging.info("Financial projection generated.")
        return projections_df

    except sqlite3.Error as e:
        logging.error(f"Error generating financial projection: {e}", exc_info=True)
        return pd.DataFrame()
    finally:
        if conn:
            conn.close()

# Example usage (for testing this module directly)
if __name__ == "__main__":
    # Ensure the database is initialized before running tests
    try:
        initialize_database()
        logging.info("Database initialized for testing db_manager.")

        # Test fetching data
        debts_df = get_table_data('Debts')
        print("\nDebts Data:")
        print(debts_df)

        accounts_df = get_table_data('Accounts')
        print("\nAccounts Data:")
        print(accounts_df)

        # Test inserting a new debt (example data)
        # new_debt_data = {
        #     'Creditor': 'Test Creditor',
        #     'OriginalAmount': 1000.0,
        #     'Amount': 1000.0,
        #     'MinimumPayment': 50.0,
        #     'SnowballPayment': 0.0,
        #     'InterestRate': 15.0,
        #     'DueDate': '2025-12-31',
        #     'Status': 'Open',
        #     'CategoryID': 1, # Assuming category ID 1 exists
        #     'AccountID': None
        # }
        # if insert_record('Debts', new_debt_data):
        #     print("\nNew debt added successfully.")
        #     print(get_table_data('Debts'))

        # Test updating debt amounts and account balances
        # (Requires some dummy data in Payments and Revenue)
        # print("\nUpdating debt amounts and account balances...")
        # update_debt_amounts_and_payments()
        # print(get_table_data('Debts'))
        # print(get_table_data('Accounts'))

        # Test generating financial projection
        # projection_df = generate_financial_projection()
        # print("\nFinancial Projection:")
        # print(projection_df)

    except Exception as e:
        logging.error(f"Test run failed: {e}", exc_info=True)
