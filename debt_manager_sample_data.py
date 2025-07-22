# debt_manager_sample_data.py
# Purpose: Populates the database with sample data for testing and demonstration.

import debt_manager_db_manager as db_manager
import logging

def populate_with_sample_data():
    """Adds a comprehensive set of sample data to the database."""
    logging.info("Populating database with sample data...")

    try:
        # 1. Add Accounts
        checking_id = db_manager.add_account_and_details({'AccountName': 'PNC Checking', 'AccountType': 'Checking', 'Balance': 2500})
        savings_id = db_manager.add_account_and_details({'AccountName': 'Ally Savings', 'AccountType': 'Savings', 'Balance': 10000})
        investment_id = db_manager.add_account_and_details({'AccountName': 'Vanguard Brokerage', 'AccountType': 'Investment', 'Balance': 50000})

        # 2. Add Debts (via Accounts)
        cc_id = db_manager.add_account_and_details(
            {'AccountName': 'Chase Sapphire', 'AccountType': 'Credit Card', 'Balance': -450.50},
            {'InterestRate': 21.99, 'MinimumPayment': 50, 'DueDate': '2025-08-15'}
        )
        loan_id = db_manager.add_account_and_details(
            {'AccountName': 'Car Loan', 'AccountType': 'Loan', 'Balance': -15000},
            {'InterestRate': 4.5, 'MinimumPayment': 350, 'DueDate': '2025-08-01'}
        )

        # 3. Add Bills (via Accounts)
        electric_id = db_manager.add_account_and_details(
            {'AccountName': 'Electric Bill', 'AccountType': 'Utilities', 'Balance': 0},
            {'EstimatedAmount': 120, 'DueDate': 20}
        )
        insurance_id = db_manager.add_account_and_details(
            {'AccountName': 'Car Insurance', 'AccountType': 'Insurance', 'Balance': 0},
            {'EstimatedAmount': 150, 'DueDate': 10}
        )

        # 4. Add Goals with linked accounts
        goal_data = {'GoalName': 'Emergency Fund', 'TargetAmount': 15000, 'TargetDate': '2026-12-31', 'Notes': '6 months of expenses'}
        db_manager.add_goal(goal_data, [savings_id]) # Link to Ally Savings

        # 5. Add Revenue with Allocations
        revenue_data = {'SourceName': 'Paycheck', 'Amount': 2000, 'DateReceived': '2025-07-15'}
        allocations = {str(checking_id): 100} # 100% to checking
        revenue_id = db_manager.add_record('Revenue', {**revenue_data, 'Allocations': json.dumps(allocations)})
        db_manager.process_revenue_allocations(revenue_id)

        # 6. Add a Payment
        payment_data = {
            'SourceAccountID': checking_id,
            'DestinationAccountID': cc_id,
            'Amount': 100,
            'PaymentDate': '2025-07-18',
            'CategoryID': db_manager.execute_query("SELECT CategoryID FROM Categories WHERE CategoryName = 'Debt Payment'", fetch='one')['CategoryID'],
            'Notes': 'Extra payment to Chase card'
        }
        db_manager.add_record('Payments', payment_data)

        logging.info("Sample data populated successfully.")

    except Exception as e:
        logging.error(f"Failed to populate sample data: {e}", exc_info=True)

if __name__ == '__main__':
    from debt_manager_db_init import initialize_database
    initialize_database()
    populate_with_sample_data()
    print("Sample data has been added to the database.")
