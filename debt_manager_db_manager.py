# debt_manager_db_manager.py
# Purpose: Manages all database interactions for the Debt Management System.
# Version: 2.3 (2025-07-22) - Added comprehensive data retrieval functions for GUI.
#          - Implemented logic for dashboard, calendar, analytics, budget, goals, and allocations.
#          - Ensured all functions return data in a GUI-friendly format (mostly pandas DataFrames).

import sqlite3
import pandas as pd
from datetime import datetime
import os
import logging
import json
from config import DB_PATH, TABLE_SCHEMAS, BUDGET_CATEGORIES

# Configure logging
LOG_DIR = os.path.join('C:\\DebtTracker', 'Logs')
LOG_FILE = os.path.join(LOG_DIR, 'DebugLog.txt')
os.makedirs(LOG_DIR, exist_ok=True)
if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s: %(message)s',
                        handlers=[logging.FileHandler(LOG_FILE, mode='a'), logging.StreamHandler()])

def get_db_connection():
    """Establishes and returns a connection to the SQLite database."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn
    except sqlite3.Error as e:
        logging.critical(f"Database connection error: {e}", exc_info=True)
        raise

def execute_query(query, params=None, fetch=None, commit=False):
    """A generic function to execute any SQL query."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query, params if params else ())

        if commit:
            conn.commit()
            last_id = cursor.lastrowid
            conn.close() # Close after commit
            return last_id

        if fetch == 'one':
            result = cursor.fetchone()
            conn.close()
            return result
        if fetch == 'all':
            result = cursor.fetchall()
            conn.close()
            return result

        # If no commit or fetch, assume the caller will handle it
        return cursor
    except sqlite3.Error as e:
        logging.error(f"Error executing query: {query} with params {params}: {e}", exc_info=True)
        if conn:
            conn.close()
        return None


def get_table_data(table_name):
    """Fetches all data from a specified table and returns a pandas DataFrame."""
    try:
        with get_db_connection() as conn:
            df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
            return df
    except (sqlite3.Error, pd.io.sql.DatabaseError) as e:
        logging.error(f"Error loading table data for {table_name}: {e}", exc_info=True)
        return pd.DataFrame()

def get_record_by_id(table_name, record_id):
    """Fetches a single record by its primary key."""
    pk_column = TABLE_SCHEMAS[table_name]['primary_key']
    query = f"SELECT * FROM {table_name} WHERE {pk_column} = ?"
    data = execute_query(query, (record_id,), fetch='one')
    return dict(data) if data else None

def add_record(table_name, data_dict):
    """Adds a new record to a table."""
    columns = ', '.join(data_dict.keys())
    placeholders = ', '.join(['?' for _ in data_dict])
    query = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
    return execute_query(query, tuple(data_dict.values()), commit=True)

def update_record(table_name, record_id, data_dict):
    """Updates an existing record in a table."""
    pk_column = TABLE_SCHEMAS[table_name]['primary_key']
    set_clause = ', '.join([f"{key} = ?" for key in data_dict])
    query = f"UPDATE {table_name} SET {set_clause} WHERE {pk_column} = ?"
    params = tuple(data_dict.values()) + (record_id,)
    execute_query(query, params, commit=True)


# --- GUI Data Retrieval Functions ---

def get_full_debt_details():
    query = """
    SELECT d.DebtID, a.AccountName, d.InterestRate, d.MinimumPayment, d.DueDate, a.Balance
    FROM Debts d JOIN Accounts a ON d.AccountID = a.AccountID
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['DebtID', 'AccountName', 'InterestRate', 'MinimumPayment', 'DueDate', 'Balance']) if data else pd.DataFrame()

def get_full_bill_details():
    query = """
    SELECT b.BillID, a.AccountName, b.EstimatedAmount, b.DueDate
    FROM Bills b JOIN Accounts a ON b.AccountID = a.AccountID
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['BillID', 'AccountName', 'EstimatedAmount', 'DueDate']) if data else pd.DataFrame()

def get_upcoming_items():
    query = """
    SELECT DueDate AS Date, AccountName AS Item, MinimumPayment AS Amount FROM Debts JOIN Accounts ON Debts.AccountID = Accounts.AccountID WHERE DueDate >= date('now')
    UNION ALL
    SELECT (strftime('%Y-%m-', date('now')) || substr('00' || DueDate, -2)) AS Date, AccountName AS Item, EstimatedAmount AS Amount FROM Bills JOIN Accounts ON Bills.AccountID = Accounts.AccountID
    ORDER BY Date
    LIMIT 10;
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['Date', 'Item', 'Amount']) if data else pd.DataFrame()

def get_goal_progress():
    query = """
    SELECT g.GoalName, g.TargetAmount, IFNULL(SUM(a.Balance), 0) as CurrentAmount
    FROM Goals g
    LEFT JOIN GoalAccountLinks gal ON g.GoalID = gal.GoalID
    LEFT JOIN Accounts a ON gal.AccountID = a.AccountID
    GROUP BY g.GoalID, g.GoalName, g.TargetAmount
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['GoalName', 'TargetAmount', 'CurrentAmount']) if data else pd.DataFrame()

def get_spending_by_category(year=datetime.now().year, month=datetime.now().month):
    query = """
    SELECT c.CategoryName, SUM(p.Amount) as TotalAmount
    FROM Payments p
    JOIN Categories c ON p.CategoryID = c.CategoryID
    WHERE CAST(strftime('%Y', p.PaymentDate) AS INTEGER) = ?
      AND CAST(strftime('%m', p.PaymentDate) AS INTEGER) = ?
      AND c.CategoryName IN ('{}')
    GROUP BY c.CategoryName
    HAVING TotalAmount > 0
    """.format("','".join(BUDGET_CATEGORIES))
    data = execute_query(query, (year, month), fetch='all')
    return pd.DataFrame(data, columns=['CategoryName', 'TotalAmount']) if data else pd.DataFrame()

def get_debt_distribution():
    query = """
    SELECT AccountName, ABS(Balance) as AbsoluteBalance
    FROM Accounts
    WHERE AccountType IN ('Credit Card', 'Loan', 'Line of Credit') AND Balance < 0
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['AccountName', 'AbsoluteBalance']) if data else pd.DataFrame()

def get_calendar_events(year, month):
    query_debts = "SELECT DueDate, AccountName FROM Debts JOIN Accounts ON Debts.AccountID = Accounts.AccountID WHERE strftime('%Y-%m', DueDate) = ?"
    query_bills = "SELECT DueDate, AccountName FROM Bills JOIN Accounts ON Bills.AccountID = Accounts.AccountID"

    month_str = f"{year}-{str(month).zfill(2)}"
    debts = execute_query(query_debts, (month_str,), fetch='all')
    bills = execute_query(query_bills, fetch='all')

    events = {}
    if debts:
        for row in debts:
            day = int(row['DueDate'].split('-')[2])
            if day not in events: events[day] = []
            events[day].append(row['AccountName'])
    if bills:
        for row in bills:
            day = int(row['DueDate'])
            if day not in events: events[day] = []
            events[day].append(row['AccountName'])
    return events

def get_budget_summary(year, month):
    query = """
    SELECT
        c.CategoryName AS Category,
        IFNULL(b.AllocatedAmount, 0) AS Allocated,
        IFNULL(p_sum.ActualAmount, 0) AS Actual
    FROM Categories c
    LEFT JOIN Budget b ON c.CategoryID = b.CategoryID
    LEFT JOIN (
        SELECT CategoryID, SUM(Amount) as ActualAmount
        FROM Payments
        WHERE CAST(strftime('%Y', PaymentDate) AS INTEGER) = ?
          AND CAST(strftime('%m', PaymentDate) AS INTEGER) = ?
        GROUP BY CategoryID
    ) p_sum ON c.CategoryID = p_sum.CategoryID
    WHERE c.CategoryName IN ('{}')
    """.format("','".join(BUDGET_CATEGORIES))
    data = execute_query(query, (year, month), fetch='all')
    return pd.DataFrame(data, columns=['Category', 'Allocated', 'Actual']) if data else pd.DataFrame()

def get_balance_history_for_account(account_name):
    query = """
    SELECT h.DateRecorded, h.Balance
    FROM BalanceHistory h
    JOIN Accounts a ON h.AccountID = a.AccountID
    WHERE a.AccountName = ?
    ORDER BY h.DateRecorded ASC
    """
    data = execute_query(query, (account_name,), fetch='all')
    return pd.DataFrame(data, columns=['DateRecorded', 'Balance']) if data else pd.DataFrame()

def get_budget_categories():
    query = "SELECT CategoryID, CategoryName FROM Categories WHERE CategoryName IN ('{}')".format("','".join(BUDGET_CATEGORIES))
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['CategoryID', 'CategoryName']) if data else pd.DataFrame()

def get_all_budgets():
    data = execute_query("SELECT CategoryID, AllocatedAmount FROM Budget", fetch='all')
    return {row['CategoryID']: row['AllocatedAmount'] for row in data} if data else {}


# --- Data Modification Functions ---

def add_account_and_details(account_data, detail_data=None):
    account_id = add_record('Accounts', account_data)
    if not account_id: return None

    account_type = account_data.get('AccountType')
    detail_data['AccountID'] = account_id
    if account_type in ['Credit Card', 'Loan', 'Line of Credit'] and detail_data:
        add_record('Debts', detail_data)
    elif account_type in ['Utilities', 'Insurance', 'Subscription'] and detail_data:
        add_record('Bills', detail_data)
    return account_id

def update_debt_details(debt_id, detail_data):
    update_record('Debts', debt_id, detail_data)

def update_bill_details(bill_id, detail_data):
    update_record('Bills', bill_id, detail_data)

def add_goal(goal_data, linked_account_ids):
    goal_id = add_record('Goals', goal_data)
    if goal_id and linked_account_ids:
        for acc_id in linked_account_ids:
            execute_query("INSERT INTO GoalAccountLinks (GoalID, AccountID) VALUES (?, ?)", (goal_id, acc_id), commit=True)

def update_goal(goal_id, goal_data, linked_account_ids):
    update_record('Goals', goal_id, goal_data)
    # Reset links and add new ones
    execute_query("DELETE FROM GoalAccountLinks WHERE GoalID = ?", (goal_id,), commit=True)
    if linked_account_ids:
        for acc_id in linked_account_ids:
            execute_query("INSERT INTO GoalAccountLinks (GoalID, AccountID) VALUES (?, ?)", (goal_id, acc_id), commit=True)

def get_linked_accounts_for_goal(goal_id):
    data = execute_query("SELECT AccountID FROM GoalAccountLinks WHERE GoalID = ?", (goal_id,), fetch='all')
    return [row['AccountID'] for row in data] if data else []

def record_all_account_balances():
    accounts = get_table_data('Accounts')
    if accounts.empty: return
    today = datetime.now().strftime("%Y-%m-%d")
    for _, row in accounts.iterrows():
        if row['Status'] == 'Active':
            # Check if a record for today already exists
            exists = execute_query("SELECT 1 FROM BalanceHistory WHERE AccountID = ? AND DateRecorded = ?", (row['AccountID'], today), fetch='one')
            if exists:
                 # Update existing record for today
                execute_query("UPDATE BalanceHistory SET Balance = ? WHERE AccountID = ? AND DateRecorded = ?", (row['Balance'], row['AccountID'], today), commit=True)
            else:
                 # Insert new record
                execute_query("INSERT INTO BalanceHistory (AccountID, DateRecorded, Balance) VALUES (?, ?, ?)", (row['AccountID'], today, row['Balance']), commit=True)

def set_budget_for_category(category_id, allocated_amount):
    exists = execute_query("SELECT BudgetID FROM Budget WHERE CategoryID = ?", (category_id,), fetch='one')
    if exists:
        execute_query("UPDATE Budget SET AllocatedAmount = ? WHERE CategoryID = ?", (allocated_amount, category_id), commit=True)
    else:
        execute_query("INSERT INTO Budget (CategoryID, AllocatedAmount) VALUES (?, ?)", (category_id, allocated_amount), commit=True)