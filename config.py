# config.py
# Purpose: Stores configuration settings and database schema definitions for the Debt Management System.
# Deploy in: C:\DebtTracker
# Version: 1.5 (2025-07-18) - Updated with complete schemas, new columns (AccountLimit, OriginalAmount, AmountPaid, CategoryID, AccountID, Notes, NextProjectedIncome, NextProjectedIncomeDate).
#          Added fields_for_new for form generation.

import os
from datetime import datetime

# --- Paths Configuration ---
BASE_DIR = "C:\\DebtTracker" # Base directory for the application
DB_DIR = os.path.join(BASE_DIR, "db")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
LOGS_DIR = os.path.join(BASE_DIR, "logs")

DB_PATH = os.path.join(DB_DIR, "debt_manager.db")
EXCEL_PATH = os.path.join(BASE_DIR, "DebtDashboard.xlsx")
REPORT_PATH = REPORTS_DIR
LOG_FILE = os.path.join(LOGS_DIR, "DebugLog.txt")
LOG_DIR = LOGS_DIR # Used for os.makedirs in orchestrator and gui

# --- Database Schema Definitions ---
# Define table schemas with:
# 'create_sql': SQL statement to create the table.
# 'db_columns': List of column names as they appear in the database.
# 'excel_columns': List of column names as they should appear in Excel headers.
# 'primary_key': The name of the primary key column.
# 'fields_for_new': List of dictionaries defining fields for 'Add New Record' forms.
#                   'name': Corresponds to 'db_columns' name.
#                   'label': Display name for the form.
#                   'type': 'text', 'real' (for numbers), 'date', 'dropdown', 'category_dropdown', 'account_selector', 'debt_account_selector'.
#                   'options': For 'dropdown' type.

TABLE_SCHEMAS = {
    "Debts": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Debts (
                DebtID INTEGER PRIMARY KEY AUTOINCREMENT,
                Creditor TEXT NOT NULL,
                OriginalAmount REAL DEFAULT 0.0,
                Amount REAL DEFAULT 0.0,
                AmountPaid REAL DEFAULT 0.0,
                MinimumPayment REAL DEFAULT 0.0,
                SnowballPayment REAL DEFAULT 0.0,
                InterestRate REAL DEFAULT 0.0,
                DueDate TEXT,
                Status TEXT,
                CategoryID INTEGER,
                AccountID INTEGER,
                FOREIGN KEY (CategoryID) REFERENCES Categories(CategoryID),
                FOREIGN KEY (AccountID) REFERENCES Accounts(AccountID)
            );
        """,
        "db_columns": [
            "DebtID", "Creditor", "OriginalAmount", "Amount", "AmountPaid",
            "MinimumPayment", "SnowballPayment", "InterestRate", "DueDate",
            "Status", "CategoryID", "AccountID"
        ],
        "excel_columns": [
            "Debt ID", "Creditor", "Original Amount", "Current Amount", "Amount Paid",
            "Minimum Payment", "Snowball Payment", "Interest Rate (%)", "Due Date",
            "Status", "Category ID", "Account ID"
        ],
        "primary_key": "DebtID",
        "fields_for_new": [
            {'name': 'Creditor', 'label': 'Creditor', 'type': 'text'},
            {'name': 'OriginalAmount', 'label': 'Original Amount', 'type': 'real'},
            {'name': 'MinimumPayment', 'label': 'Minimum Payment', 'type': 'real'},
            {'name': 'SnowballPayment', 'label': 'Snowball Payment', 'type': 'real'},
            {'name': 'InterestRate', 'label': 'Interest Rate (%)', 'type': 'real'},
            {'name': 'DueDate', 'label': 'Due Date', 'type': 'date'},
            {'name': 'Status', 'label': 'Status', 'type': 'dropdown', 'options': ['Open', 'Paid', 'Paid Off', 'Closed', 'Defaulted']},
            {'name': 'CategoryID', 'label': 'Category', 'type': 'category_dropdown'}, # Links to Categories table
            {'name': 'AccountID', 'label': 'Linked Account', 'type': 'account_selector'} # Links to Accounts table
        ]
    },
    "Accounts": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Accounts (
                AccountID INTEGER PRIMARY KEY AUTOINCREMENT,
                AccountName TEXT NOT NULL,
                AccountType TEXT,
                StartingBalance REAL DEFAULT 0.0,
                Balance REAL DEFAULT 0.0,
                PreviousBalance REAL DEFAULT 0.0,
                AccountLimit REAL DEFAULT 0.0,
                Status TEXT
            );
        """,
        "db_columns": [
            "AccountID", "AccountName", "AccountType", "StartingBalance",
            "Balance", "PreviousBalance", "AccountLimit", "Status"
        ],
        "excel_columns": [
            "Account ID", "Account Name", "Account Type", "Starting Balance",
            "Current Balance", "Previous Balance", "Account Limit", "Status"
        ],
        "primary_key": "AccountID",
        "fields_for_new": [
            {'name': 'AccountName', 'label': 'Account Name', 'type': 'text'},
            {'name': 'AccountType', 'label': 'Account Type', 'type': 'dropdown', 'options': ['Checking', 'Savings', 'Credit Card', 'Loan', 'Investment', 'Other']},
            {'name': 'StartingBalance', 'label': 'Starting Balance', 'type': 'real'},
            {'name': 'AccountLimit', 'label': 'Account Limit (for Credit/Loan)', 'type': 'real'},
            {'name': 'Status', 'label': 'Status', 'type': 'dropdown', 'options': ['Open', 'Closed', 'Active', 'Inactive']}
        ]
    },
    "Payments": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Payments (
                PaymentID INTEGER PRIMARY KEY AUTOINCREMENT,
                DebtID INTEGER,
                AccountID INTEGER,
                Amount REAL NOT NULL,
                PaymentDate TEXT,
                PaymentMethod TEXT,
                Category TEXT,
                Notes TEXT,
                FOREIGN KEY (DebtID) REFERENCES Debts(DebtID),
                FOREIGN KEY (AccountID) REFERENCES Accounts(AccountID)
            );
        """,
        "db_columns": [
            "PaymentID", "DebtID", "AccountID", "Amount", "PaymentDate",
            "PaymentMethod", "Category", "Notes"
        ],
        "excel_columns": [
            "Payment ID", "Debt ID", "Account ID", "Amount", "Payment Date",
            "Payment Method", "Category", "Notes"
        ],
        "primary_key": "PaymentID",
        "fields_for_new": [
            {'name': 'DebtID', 'label': 'Linked Debt', 'type': 'debt_account_selector'}, # Can link to Debt or Account
            {'name': 'AccountID', 'label': 'Payment From Account', 'type': 'account_selector'}, # Links to Accounts table
            {'name': 'Amount', 'label': 'Amount', 'type': 'real'},
            {'name': 'PaymentDate', 'label': 'Payment Date', 'type': 'date'},
            {'name': 'PaymentMethod', 'label': 'Payment Method', 'type': 'dropdown', 'options': ['Direct Debit', 'Credit Card', 'Bank Transfer', 'Cheque', 'Cash', 'Other']},
            {'name': 'Category', 'label': 'Category', 'type': 'category_dropdown'}, # Links to Categories table
            {'name': 'Notes', 'label': 'Notes', 'type': 'text'}
        ]
    },
    "Goals": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Goals (
                GoalID INTEGER PRIMARY KEY AUTOINCREMENT,
                GoalName TEXT NOT NULL,
                TargetAmount REAL DEFAULT 0.0,
                CurrentAmount REAL DEFAULT 0.0,
                TargetDate TEXT,
                Status TEXT,
                Notes TEXT
            );
        """,
        "db_columns": [
            "GoalID", "GoalName", "TargetAmount", "CurrentAmount",
            "TargetDate", "Status", "Notes"
        ],
        "excel_columns": [
            "Goal ID", "Goal Name", "Target Amount", "Current Amount",
            "Target Date", "Status", "Notes"
        ],
        "primary_key": "GoalID",
        "fields_for_new": [
            {'name': 'GoalName', 'label': 'Goal Name', 'type': 'text'},
            {'name': 'TargetAmount', 'label': 'Target Amount', 'type': 'real'},
            {'name': 'CurrentAmount', 'label': 'Current Amount', 'type': 'real'},
            {'name': 'TargetDate', 'label': 'Target Date', 'type': 'date'},
            {'name': 'Status', 'label': 'Status', 'type': 'dropdown', 'options': ['Planned', 'In Progress', 'Completed', 'On Hold', 'Cancelled']},
            {'name': 'Notes', 'label': 'Notes', 'type': 'text'}
        ]
    },
    "Assets": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Assets (
                AssetID INTEGER PRIMARY KEY AUTOINCREMENT,
                AssetName TEXT NOT NULL,
                Value REAL DEFAULT 0.0,
                Category TEXT,
                AssetStatus TEXT,
                PurchaseDate TEXT,
                Notes TEXT
            );
        """,
        "db_columns": [
            "AssetID", "AssetName", "Value", "Category", "AssetStatus", "PurchaseDate", "Notes"
        ],
        "excel_columns": [
            "Asset ID", "Asset Name", "Value", "Category", "Status", "Purchase Date", "Notes"
        ],
        "primary_key": "AssetID",
        "fields_for_new": [
            {'name': 'AssetName', 'label': 'Asset Name', 'type': 'text'},
            {'name': 'Value', 'label': 'Current Value', 'type': 'real'},
            {'name': 'Category', 'label': 'Category', 'type': 'category_dropdown'},
            {'name': 'AssetStatus', 'label': 'Status', 'type': 'dropdown', 'options': ['Active', 'Sold', 'Disposed', 'Liquidated']},
            {'name': 'PurchaseDate', 'label': 'Purchase Date', 'type': 'date'},
            {'name': 'Notes', 'label': 'Notes', 'type': 'text'}
        ]
    },
    "Revenue": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Revenue (
                RevenueID INTEGER PRIMARY KEY AUTOINCREMENT,
                Amount REAL NOT NULL,
                DateReceived TEXT,
                Source TEXT,
                AllocatedTo INTEGER,
                AllocationPercentage REAL DEFAULT 0.0,
                AllocationType TEXT,
                NextProjectedIncome REAL DEFAULT 0.0,
                NextProjectedIncomeDate TEXT
            );
        """,
        "db_columns": [
            "RevenueID", "Amount", "DateReceived", "Source", "AllocatedTo",
            "AllocationPercentage", "AllocationType", "NextProjectedIncome", "NextProjectedIncomeDate"
        ],
        "excel_columns": [
            "Revenue ID", "Amount", "Date Received", "Source", "Allocated To ID",
            "Allocation Percentage", "Allocation Type", "Next Projected Income", "Next Projected Income Date"
        ],
        "primary_key": "RevenueID",
        "fields_for_new": [
            {'name': 'Amount', 'label': 'Amount', 'type': 'real'},
            {'name': 'DateReceived', 'label': 'Date Received', 'type': 'date'},
            {'name': 'Source', 'label': 'Source', 'type': 'text'},
            {'name': 'AllocatedTo', 'label': 'Allocated To (Account/Goal ID)', 'type': 'debt_account_selector'}, # Can be AccountID or GoalID
            {'name': 'AllocationPercentage', 'label': 'Allocation Percentage (%)', 'type': 'real'},
            {'name': 'AllocationType', 'label': 'Allocation Type', 'type': 'dropdown', 'options': ['Account', 'Goal', 'Category-Based', 'Unallocated']},
            {'name': 'NextProjectedIncome', 'label': 'Next Projected Income', 'type': 'real'},
            {'name': 'NextProjectedIncomeDate', 'label': 'Next Projected Income Date', 'type': 'date'}
        ]
    },
    "Categories": {
        "create_sql": """
            CREATE TABLE IF NOT EXISTS Categories (
                CategoryID INTEGER PRIMARY KEY AUTOINCREMENT,
                CategoryName TEXT NOT NULL UNIQUE
            );
        """,
        "db_columns": [
            "CategoryID", "CategoryName"
        ],
        "excel_columns": [
            "Category ID", "Category Name"
        ],
        "primary_key": "CategoryID",
        "fields_for_new": [
            {'name': 'CategoryName', 'label': 'Category Name', 'type': 'text'}
        ]
    }
}

# --- Predefined Categories ---
PREDEFINED_CATEGORIES = [
    {"CategoryName": "Housing"},
    {"CategoryName": "Utilities"},
    {"CategoryName": "Groceries"},
    {"CategoryName": "Transportation"},
    {"CategoryName": "Dining Out"},
    {"CategoryName": "Entertainment"},
    {"CategoryName": "Health & Fitness"},
    {"CategoryName": "Shopping"},
    {"CategoryName": "Education"},
    {"CategoryName": "Personal Care"},
    {"CategoryName": "Miscellaneous"},
    {"CategoryName": "Income"},
    {"CategoryName": "Savings"},
    {"CategoryName": "Investment"},
    {"CategoryName": "Debt Payment"},
    {"CategoryName": "Credit Card"}, # For credit card debts
    {"CategoryName": "Loan"},        # For loan debts
    {"CategoryName": "Bills"},       # For general bills
    {"CategoryName": "Collection"},  # For collection accounts
    {"CategoryName": "Emergency Fund"},
    {"CategoryName": "Retirement"}
]
