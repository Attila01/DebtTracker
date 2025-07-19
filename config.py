# config.py
# Purpose: Centralized configuration for the Debt Management System.
#          Defines paths, database schemas, and now CSV file mappings.
# Deploy in: C:\DebtTracker
# Version: 1.5 (2025-07-19) - Added AccountID to Payments schema for linking payments to source accounts.
#                            Updated paths and schema references accordingly.

import os

# Base directory for the application
BASE_DIR = 'C:\\DebtTracker'

# Database configuration
DB_DIR = os.path.join(BASE_DIR, 'db')
DB_PATH = os.path.join(DB_DIR, 'debt_manager.db') # SQLite database path

# CSV configuration
CSV_DIR = os.path.join(BASE_DIR, 'csv_data') # New directory for CSV files
# Note: Each table will have its own CSV file (e.g., Debts.csv, Accounts.csv)

# Logging configuration
LOG_DIR = os.path.join(BASE_DIR, 'Logs')
LOG_FILE = os.path.join(LOG_DIR, 'DebugLog.txt') # Consolidated log file

# Define table schemas for SQLite and CSV mapping
# 'columns': SQLite table schema (name, type, primary_key, nullable)
# 'csv_columns': Order and names of columns as they should appear in CSV
# 'gui_fields': Details for GUI form generation (name, type, options for comboboxes)
TABLE_SCHEMAS = {
    'Debts': {
        'columns': [
            {'name': 'DebtID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'Creditor', 'type': 'TEXT', 'nullable': False},
            {'name': 'OriginalAmount', 'type': 'REAL', 'nullable': False}, # Added for tracking original debt
            {'name': 'Amount', 'type': 'REAL', 'nullable': False},
            {'name': 'AmountPaid', 'type': 'REAL', 'nullable': True, 'default': 0.0}, # Track total paid
            {'name': 'MinimumPayment', 'type': 'REAL', 'nullable': True},
            {'name': 'SnowballPayment', 'type': 'REAL', 'nullable': True},
            {'name': 'InterestRate', 'type': 'REAL', 'nullable': True},
            {'name': 'DueDate', 'type': 'TEXT', 'nullable': True}, # Stored as ISO 8601 string 'YYYY-MM-DD'
            {'name': 'Status', 'type': 'TEXT', 'nullable': True},
            {'name': 'CategoryID', 'type': 'INTEGER', 'nullable': True}, # Foreign key to Categories
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': True} # Foreign key to Accounts (for linked accounts)
        ],
        'csv_columns': [ # Renamed from excel_columns
            'DebtID', 'Creditor', 'OriginalAmount', 'Amount', 'AmountPaid', 'MinimumPayment',
            'SnowballPayment', 'InterestRate', 'DueDate', 'Status', 'CategoryID', 'AccountID'
        ],
        'gui_fields': [
            {'name': 'Creditor', 'type': 'text'},
            {'name': 'OriginalAmount', 'type': 'decimal'},
            {'name': 'Amount', 'type': 'decimal'},
            {'name': 'MinimumPayment', 'type': 'decimal'},
            {'name': 'SnowballPayment', 'type': 'decimal'},
            {'name': 'InterestRate', 'type': 'decimal'},
            {'name': 'DueDate', 'type': 'date'},
            {'name': 'Status', 'type': 'combo', 'options': ['Open', 'Closed', 'Current', 'In Collection', 'Paid Off']},
            {'name': 'Category', 'type': 'combo', 'source_table': 'Categories', 'source_display_col': 'CategoryName', 'source_value_col': 'CategoryID'},
            {'name': 'Account', 'type': 'combo', 'source_table': 'Accounts', 'source_display_col': 'AccountName', 'source_value_col': 'AccountID'}
        ],
        'primary_key': 'DebtID'
    },
    'Accounts': {
        'columns': [
            {'name': 'AccountID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'AccountName', 'type': 'TEXT', 'nullable': False},
            {'name': 'Balance', 'type': 'REAL', 'nullable': False},
            {'name': 'AccountType', 'type': 'TEXT', 'nullable': True},
            {'name': 'Status', 'type': 'TEXT', 'nullable': True},
            {'name': 'AccountLimit', 'type': 'REAL', 'nullable': True, 'default': 0.0}, # For credit cards/lines of credit
            {'name': 'InitialBalance', 'type': 'REAL', 'nullable': True, 'default': 0.0} # Added for balance tracking
        ],
        'csv_columns': [ # Renamed from excel_columns
            'AccountID', 'AccountName', 'Balance', 'AccountType', 'Status', 'AccountLimit', 'InitialBalance'
        ],
        'gui_fields': [
            {'name': 'AccountName', 'type': 'text'},
            {'name': 'Balance', 'type': 'decimal'},
            {'name': 'AccountType', 'type': 'combo', 'options': ['Checking', 'Savings', 'Credit Card', 'Loan', 'Investment']},
            {'name': 'Status', 'type': 'combo', 'options': ['Open', 'Closed', 'Current', 'Active', 'Inactive']},
            {'name': 'AccountLimit', 'type': 'decimal'},
            {'name': 'InitialBalance', 'type': 'decimal'}
        ],
        'primary_key': 'AccountID'
    },
    'Payments': {
        'columns': [
            {'name': 'PaymentID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'DebtID', 'type': 'INTEGER', 'nullable': True}, # Can be null if payment is not for a specific debt (e.g., general expense)
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': True}, # NEW: Link to the account from which payment was made
            {'name': 'Amount', 'type': 'REAL', 'nullable': False},
            {'name': 'PaymentDate', 'type': 'TEXT', 'nullable': False}, # Stored as ISO 8601 string 'YYYY-MM-DD'
            {'name': 'PaymentMethod', 'type': 'TEXT', 'nullable': True},
            {'name': 'CategoryID', 'type': 'INTEGER', 'nullable': True}, # Added this column for linking to Categories
            {'name': 'Notes', 'type': 'TEXT', 'nullable': True} # Added notes field
        ],
        'csv_columns': [ # Renamed from excel_columns
            'PaymentID', 'DebtID', 'AccountID', 'Amount', 'PaymentDate', 'PaymentMethod', 'CategoryID', 'Notes' # Added AccountID
        ],
        'gui_fields': [
            {'name': 'Debt', 'type': 'combo', 'source_table': 'Debts', 'source_display_col': 'Creditor', 'source_value_col': 'DebtID', 'allow_none': True}, # Allow 'None' for non-debt payments
            {'name': 'Account', 'type': 'combo', 'source_table': 'Accounts', 'source_display_col': 'AccountName', 'source_value_col': 'AccountID', 'allow_none': False}, # NEW: Payment source account, typically not None
            {'name': 'Amount', 'type': 'decimal'},
            {'name': 'PaymentDate', 'type': 'date'},
            {'name': 'PaymentMethod', 'type': 'text'},
            {'name': 'Category', 'type': 'combo', 'source_table': 'Categories', 'source_display_col': 'CategoryName', 'source_value_col': 'CategoryID', 'allow_none': True}, # Link to CategoryID
            {'name': 'Notes', 'type': 'text'}
        ],
        'primary_key': 'PaymentID'
    },
    'Goals': {
        'columns': [
            {'name': 'GoalID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'GoalName', 'type': 'TEXT', 'nullable': False},
            {'name': 'TargetAmount', 'type': 'REAL', 'nullable': False},
            {'name': 'CurrentAmount', 'type': 'REAL', 'nullable': True, 'default': 0.0},
            {'name': 'TargetDate', 'type': 'TEXT', 'nullable': True}, # Stored as ISO 8601 string 'YYYY-MM-DD'
            {'name': 'Status', 'type': 'TEXT', 'nullable': True},
            {'name': 'Notes', 'type': 'TEXT', 'nullable': True},
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': True} # Added AccountID for linking to accounts
        ],
        'csv_columns': [ # Renamed from excel_columns
            'GoalID', 'GoalName', 'TargetAmount', 'CurrentAmount', 'TargetDate', 'Status', 'Notes', 'AccountID'
        ],
        'gui_fields': [
            {'name': 'GoalName', 'type': 'text'},
            {'name': 'TargetAmount', 'type': 'decimal'},
            {'name': 'CurrentAmount', 'type': 'decimal'},
            {'name': 'TargetDate', 'type': 'date'},
            {'name': 'Status', 'type': 'combo', 'options': ['Planned', 'In Progress', 'Completed', 'On Hold']},
            {'name': 'Notes', 'type': 'text'},
            {'name': 'Account', 'type': 'combo', 'source_table': 'Accounts', 'source_display_col': 'AccountName', 'source_value_col': 'AccountID', 'allow_none': True} # Link to AccountID
        ],
        'primary_key': 'GoalID'
    },
    'Assets': {
        'columns': [
            {'name': 'AssetID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'AssetName', 'type': 'TEXT', 'nullable': False},
            {'name': 'Value', 'type': 'REAL', 'nullable': False},
            {'name': 'Category', 'type': 'TEXT', 'nullable': True}, # Category name from Categories table
            {'name': 'PurchaseDate', 'type': 'TEXT', 'nullable': True}, # Stored as ISO 8601 string 'YYYY-MM-DD'
            {'name': 'Status', 'type': 'TEXT', 'nullable': True},
            {'name': 'Notes', 'type': 'TEXT', 'nullable': True}
        ],
        'csv_columns': [ # Renamed from excel_columns
            'AssetID', 'AssetName', 'Value', 'Category', 'PurchaseDate', 'Status', 'Notes'
        ],
        'gui_fields': [
            {'name': 'AssetName', 'type': 'text'},
            {'name': 'Value', 'type': 'decimal'},
            {'name': 'Category', 'type': 'combo', 'source_table': 'Categories', 'source_display_col': 'CategoryName', 'source_value_col': 'CategoryName'},
            {'name': 'PurchaseDate', 'type': 'date'},
            {'name': 'Status', 'type': 'combo', 'options': ['Active', 'Sold', 'Disposed']},
            {'name': 'Notes', 'type': 'text'}
        ],
        'primary_key': 'AssetID'
    },
    'Revenue': {
        'columns': [
            {'name': 'RevenueID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'Amount', 'type': 'REAL', 'nullable': False},
            {'name': 'DateReceived', 'type': 'TEXT', 'nullable': False}, # Stored as ISO 8601 string 'YYYY-MM-DD'
            {'name': 'Source', 'type': 'TEXT', 'nullable': True},
            {'name': 'AllocatedTo', 'type': 'INTEGER', 'nullable': True}, # Foreign key to AccountID or DebtID
            {'name': 'AllocationType', 'type': 'TEXT', 'nullable': True}, # 'Account', 'Debt', 'Other'
            {'name': 'NextProjectedIncome', 'type': 'REAL', 'nullable': True}, # For recurring income tracking
            {'name': 'NextProjectedIncomeDate', 'type': 'TEXT', 'nullable': True}, # Stored as ISO 8601 string 'YYYY-MM-DD'
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': True} # Added AccountID for direct account linking
        ],
        'csv_columns': [ # Renamed from excel_columns
            'RevenueID', 'Amount', 'DateReceived', 'Source', 'AllocatedTo', 'AllocationType',
            'NextProjectedIncome', 'NextProjectedIncomeDate', 'AccountID'
        ],
        'gui_fields': [
            {'name': 'Amount', 'type': 'decimal'},
            {'name': 'DateReceived', 'type': 'date'},
            {'name': 'Source', 'type': 'text'},
            # Updated AllocatedTo to also show AccountName for better context in the GUI
            {'name': 'AllocatedTo', 'type': 'combo', 'source_table': ['Accounts', 'Debts'], 'source_display_col': ['AccountName', 'Creditor'], 'source_value_col': ['AccountID', 'DebtID'], 'allow_none': True},
            {'name': 'AllocationType', 'type': 'combo', 'options': ['Account', 'Debt', 'Other']},
            {'name': 'NextProjectedIncome', 'type': 'decimal'},
            {'name': 'NextProjectedIncomeDate', 'type': 'date'},
            {'name': 'Account', 'type': 'combo', 'source_table': 'Accounts', 'source_display_col': 'AccountName', 'source_value_col': 'AccountID', 'allow_none': True} # Direct link to Account for convenience
        ],
        'primary_key': 'RevenueID'
    },
    'Categories': {
        'columns': [
            {'name': 'CategoryID', 'type': 'INTEGER', 'primary_key': True, 'nullable': False, 'autoincrement': True},
            {'name': 'CategoryName', 'type': 'TEXT', 'nullable': False, 'unique': True}
        ],
        'csv_columns': [ # Renamed from excel_columns
            'CategoryID', 'CategoryName'
        ],
        'gui_fields': [
            {'name': 'CategoryName', 'type': 'text'}
        ],
        'primary_key': 'CategoryID'
    }
}

# Predefined categories for initial database setup
PREDEFINED_CATEGORIES = [
    "Housing", "Utilities", "Groceries", "Transportation", "Healthcare",
    "Insurance", "Debt Payment", "Savings", "Investments", "Education",
    "Entertainment", "Dining Out", "Shopping", "Personal Care", "Gifts/Donations",
    "Miscellaneous", "Salary", "Freelance Income", "Bonus", "Refund", "Interest Income"
]
