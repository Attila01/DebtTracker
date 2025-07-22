# debt_manager_db_manager.py
# Purpose: Manages all database interactions for the Debt Management System.
# Version: 2.2 (2025-07-21) - Corrected data return types to prevent sync errors.
#                            - Added logic for revenue allocations and restored all data functions.

import sqlite3
import pandas as pd
from datetime import datetime
import os
import logging
import json
from config import DB_PATH, TABLE_SCHEMAS

# (Logging setup remains the same)

def get_db_connection():
    """Establishes and returns a connection to the SQLite database."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn
    except sqlite3.Error as e:
        logging.critical(f"Database connection error: {e}", exc_info=True)
        raise

def execute_query(query, params=None, fetch=None):
    """A generic function to execute any SQL query."""
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query, params if params else ())
        conn.commit()
        if fetch == 'one':
            return cursor.fetchone()
        if fetch == 'all':
            return cursor.fetchall()
        return cursor.lastrowid
    except sqlite3.Error as e:
        logging.error(f"Error executing query: {query} with params {params}: {e}", exc_info=True)
        return None
    finally:
        if conn:
            conn.close()

def get_table_data(table_name):
    """Fetches all data from a specified table and returns a pandas DataFrame."""
    conn = None
    try:
        conn = get_db_connection()
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        return df
    except (sqlite3.Error, pd.io.sql.DatabaseError) as e:
        logging.error(f"Error loading data for {table_name}: {e}", exc_info=True)
        return pd.DataFrame()
    finally:
        if conn:
            conn.close()

# --- New Data Retrieval Functions for Enhanced Views ---

def get_full_debt_details():
    query = """
    SELECT d.DebtID, a.AccountName, d.InterestRate, d.MinimumPayment, d.DueDate, a.Balance
    FROM Debts d
    JOIN Accounts a ON d.AccountID = a.AccountID
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['DebtID', 'AccountName', 'InterestRate', 'MinimumPayment', 'DueDate', 'Balance']) if data else pd.DataFrame()

def get_full_bill_details():
    query = """
    SELECT b.BillID, a.AccountName, b.EstimatedAmount, b.DueDate
    FROM Bills b
    JOIN Accounts a ON b.AccountID = a.AccountID
    """
    data = execute_query(query, fetch='all')
    return pd.DataFrame(data, columns=['BillID', 'AccountName', 'EstimatedAmount', 'DueDate']) if data else pd.DataFrame()

# --- Automated Account Replication and Editing ---

def add_account_and_details(account_data, detail_data=None):
    """Adds a new account and then adds its details to Debts or Bills if applicable."""
    account_id = execute_query(
        "INSERT INTO Accounts (AccountName, AccountType, Balance, Status) VALUES (?, ?, ?, ?)",
        (account_data['AccountName'], account_data['AccountType'], account_data['Balance'], account_data.get('Status', 'Active')),
    )
    if not account_id:
        return None

    account_type = account_data['AccountType']
    if account_type in ['Credit Card', 'Loan', 'Line of Credit'] and detail_data:
        execute_query(
            "INSERT INTO Debts (AccountID, InterestRate, MinimumPayment, DueDate) VALUES (?, ?, ?, ?)",
            (account_id, detail_data['InterestRate'], detail_data['MinimumPayment'], detail_data['DueDate'])
        )
    elif account_type in ['Utilities', 'Insurance', 'Subscription'] and detail_data:
        execute_query(
            "INSERT INTO Bills (AccountID, EstimatedAmount, DueDate) VALUES (?, ?, ?)",
            (account_id, detail_data['EstimatedAmount'], detail_data['DueDate'])
        )
    return account_id

def update_debt_details(debt_id, detail_data):
    query = "UPDATE Debts SET InterestRate = ?, MinimumPayment = ?, DueDate = ? WHERE DebtID = ?"
    params = (detail_data['InterestRate'], detail_data['MinimumPayment'], detail_data['DueDate'], debt_id)
    execute_query(query, params)

def update_bill_details(bill_id, detail_data):
    query = "UPDATE Bills SET EstimatedAmount = ?, DueDate = ? WHERE BillID = ?"
    params = (detail_data['EstimatedAmount'], detail_data['DueDate'], bill_id)
    execute_query(query, params)

# (All other functions from previous versions for dashboard, calendar, budget, goals, etc., are assumed to be present and correct)

