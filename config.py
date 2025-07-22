# config.py
# Purpose: Centralized configuration for the Debt Management System.
# Version: 2.4 (2025-07-21) - Corrected Debts schema by removing the redundant 'Creditor' column.

import os

BASE_DIR = 'C:\\DebtTracker'
DB_DIR = os.path.join(BASE_DIR, 'db')
DB_PATH = os.path.join(DB_DIR, 'debt_manager.db')
CSV_DIR = os.path.join(BASE_DIR, 'csv_data')
LOG_DIR = os.path.join(BASE_DIR, 'Logs')
LOG_FILE = os.path.join(LOG_DIR, 'DebugLog.txt')

TABLE_SCHEMAS = {
    'Accounts': {
        'columns': [
            {'name': 'AccountID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'AccountName', 'type': 'TEXT', 'nullable': False, 'unique': True},
            {'name': 'AccountType', 'type': 'TEXT', 'nullable': False},
            {'name': 'Balance', 'type': 'REAL', 'nullable': False, 'default': 0.0},
            {'name': 'Status', 'type': 'TEXT', 'nullable': True, 'default': 'Active'},
        ],
        'csv_columns': ['AccountID', 'AccountName', 'AccountType', 'Balance', 'Status'],
        'gui_fields': [
            {'name': 'AccountName', 'type': 'text'},
            {'name': 'AccountType', 'type': 'combo', 'options': [
                'Checking', 'Savings', 'Investment', 'Cash',
                'Credit Card', 'Loan', 'Line of Credit',
                'Utilities', 'Insurance', 'Subscription'
            ]},
            {'name': 'Balance', 'type': 'decimal'},
            {'name': 'Status', 'type': 'combo', 'options': ['Active', 'Inactive', 'Closed']},
        ],
        'primary_key': 'AccountID'
    },
    'Debts': {
        'columns': [
            {'name': 'DebtID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': False},
            {'name': 'InterestRate', 'type': 'REAL', 'nullable': True, 'default': 0.0},
            {'name': 'MinimumPayment', 'type': 'REAL', 'nullable': True, 'default': 0.0},
            {'name': 'DueDate', 'type': 'TEXT', 'nullable': True},
        ],
        'csv_columns': ['DebtID', 'AccountID', 'InterestRate', 'MinimumPayment', 'DueDate'],
        'gui_fields': [
            {'name': 'InterestRate', 'type': 'decimal'},
            {'name': 'MinimumPayment', 'type': 'decimal'},
            {'name': 'DueDate', 'type': 'date'},
        ],
        'primary_key': 'DebtID'
    },
    'Bills': {
        'columns': [
            {'name': 'BillID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': False},
            {'name': 'EstimatedAmount', 'type': 'REAL', 'nullable': True, 'default': 0.0},
            {'name': 'DueDate', 'type': 'INTEGER', 'nullable': True},
        ],
        'csv_columns': ['BillID', 'AccountID', 'EstimatedAmount', 'DueDate'],
        'gui_fields': [
            {'name': 'EstimatedAmount', 'type': 'decimal'},
            {'name': 'DueDate', 'type': 'integer'},
        ],
        'primary_key': 'BillID'
    },
    'Revenue': {
        'columns': [
            {'name': 'RevenueID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'SourceName', 'type': 'TEXT', 'nullable': False},
            {'name': 'Amount', 'type': 'REAL', 'nullable': False},
            {'name': 'DateReceived', 'type': 'TEXT', 'nullable': False},
            {'name': 'Allocations', 'type': 'TEXT', 'nullable': True}
        ],
        'csv_columns': ['RevenueID', 'SourceName', 'Amount', 'DateReceived', 'Allocations'],
        'gui_fields': [
            {'name': 'SourceName', 'type': 'text'},
            {'name': 'Amount', 'type': 'decimal'},
            {'name': 'DateReceived', 'type': 'date'},
            {'name': 'Allocations', 'type': 'allocations'}
        ],
        'primary_key': 'RevenueID'
    },
    'Payments': {
        'columns': [
            {'name': 'PaymentID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'SourceAccountID', 'type': 'INTEGER', 'nullable': False},
            {'name': 'DestinationAccountID', 'type': 'INTEGER', 'nullable': True},
            {'name': 'Amount', 'type': 'REAL', 'nullable': False},
            {'name': 'PaymentDate', 'type': 'TEXT', 'nullable': False},
            {'name': 'CategoryID', 'type': 'INTEGER', 'nullable': False},
            {'name': 'Notes', 'type': 'TEXT', 'nullable': True}
        ],
        'csv_columns': ['PaymentID', 'SourceAccountID', 'DestinationAccountID', 'Amount', 'PaymentDate', 'CategoryID', 'Notes'],
        'gui_fields': [
            {'name': 'Source Account', 'type': 'combo', 'source_table': 'Accounts'},
            {'name': 'Destination Account', 'type': 'combo', 'source_table': 'Accounts', 'allow_none': True},
            {'name': 'Amount', 'type': 'decimal'},
            {'name': 'PaymentDate', 'type': 'date'},
            {'name': 'Category', 'type': 'combo', 'source_table': 'Categories'},
            {'name': 'Notes', 'type': 'text'}
        ],
        'primary_key': 'PaymentID'
    },
    'Budget': {
        'columns': [
            {'name': 'BudgetID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'CategoryID', 'type': 'INTEGER', 'nullable': False, 'unique': True},
            {'name': 'AllocatedAmount', 'type': 'REAL', 'nullable': False, 'default': 0.0},
        ],
        'csv_columns': ['BudgetID', 'CategoryID', 'AllocatedAmount'],
        'gui_fields': [
            {'name': 'Category', 'type': 'combo', 'source_table': 'Categories'},
            {'name': 'AllocatedAmount', 'type': 'decimal'}
        ],
        'primary_key': 'BudgetID'
    },
    'BalanceHistory': {
        'columns': [
            {'name': 'HistoryID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': False},
            {'name': 'DateRecorded', 'type': 'TEXT', 'nullable': False},
            {'name': 'Balance', 'type': 'REAL', 'nullable': False}
        ],
        'csv_columns': ['HistoryID', 'AccountID', 'DateRecorded', 'Balance'],
        'gui_fields': [],
        'primary_key': 'HistoryID'
    },
    'Goals': {
        'columns': [
            {'name': 'GoalID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'GoalName', 'type': 'TEXT', 'nullable': False},
            {'name': 'TargetAmount', 'type': 'REAL', 'nullable': False},
            {'name': 'TargetDate', 'type': 'TEXT', 'nullable': True},
            {'name': 'Notes', 'type': 'TEXT', 'nullable': True},
        ],
        'csv_columns': ['GoalID', 'GoalName', 'TargetAmount', 'TargetDate', 'Notes'],
        'gui_fields': [
            {'name': 'GoalName', 'type': 'text'},
            {'name': 'TargetAmount', 'type': 'decimal'},
            {'name': 'TargetDate', 'type': 'date'},
            {'name': 'Notes', 'type': 'text'}
        ],
        'primary_key': 'GoalID'
    },
    'GoalAccountLinks': {
        'columns': [
            {'name': 'LinkID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'GoalID', 'type': 'INTEGER', 'nullable': False},
            {'name': 'AccountID', 'type': 'INTEGER', 'nullable': False},
        ],
        'csv_columns': ['LinkID', 'GoalID', 'AccountID'],
        'gui_fields': [],
        'primary_key': 'LinkID'
    },
    'Categories': {
        'columns': [
            {'name': 'CategoryID', 'type': 'INTEGER', 'primary_key': True, 'autoincrement': True},
            {'name': 'CategoryName', 'type': 'TEXT', 'nullable': False, 'unique': True}
        ],
        'csv_columns': ['CategoryID', 'CategoryName'],
        'gui_fields': [{'name': 'CategoryName', 'type': 'text'}],
        'primary_key': 'CategoryID'
    }
}

PREDEFINED_CATEGORIES = [
    "Housing", "Utilities", "Groceries", "Transportation", "Healthcare",
    "Insurance", "Entertainment", "Shopping", "Gifts/Donations",
    "Salary", "Freelance Income", "Investment Income", "Debt Payment", "Savings Transfer", "Miscellaneous"
]

BUDGET_CATEGORIES = [
    "Housing", "Utilities", "Groceries", "Transportation", "Healthcare",
    "Insurance", "Entertainment", "Shopping", "Gifts/Donations", "Miscellaneous"
]
