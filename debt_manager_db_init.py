# debt_manager_db_init.py
# Purpose: Handles the initial creation of the SQLite database and its tables,
#          including inserting predefined categories.
# Deploy in: C:\DebtTracker
# Version: 1.3 (2025-07-18) - Updated table creation SQL to include new columns
#          (AccountLimit, OriginalAmount, AmountPaid, CategoryID, AccountID, Notes, NextProjectedIncome, NextProjectedIncomeDate)
#          and improved logging for database initialization.

import sqlite3
import os
import logging
from config import DB_PATH, DB_DIR, TABLE_SCHEMAS, PREDEFINED_CATEGORIES, LOG_FILE, LOG_DIR

# Ensure log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

# Configure logging (if not already configured by orchestrator)
if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s: %(message)s',
                        handlers=[
                            logging.FileHandler(LOG_FILE, mode='a'),
                            logging.StreamHandler()
                        ])

def initialize_database():
    """
    Initializes the SQLite database:
    - Creates the database file if it doesn't exist.
    - Creates all necessary tables based on TABLE_SCHEMAS if they don't exist.
    - Inserts predefined categories if the Categories table is empty.
    """
    logging.info("Starting database initialization process.")

    # Ensure the database directory exists
    os.makedirs(DB_DIR, exist_ok=True)

    conn = None
    try:
        # Check if the database file exists and is a valid SQLite database
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                # Attempt a simple query to verify it's a valid SQLite DB
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                logging.info(f"Existing database file '{DB_PATH}' is a valid SQLite database.")
            except sqlite3.DatabaseError as e:
                logging.critical(f"Critical database initialization error: {e}")
                logging.critical(f"Existing file at '{DB_PATH}' is not a valid SQLite database. Please delete or move it if you want to create a new one.")
                raise # Re-raise to stop execution if DB is corrupt
        else:
            logging.info(f"Database file not found. Attempting to create: {DB_PATH}")
            conn = sqlite3.connect(DB_PATH) # This creates the file
            logging.info(f"Database file created successfully: {DB_PATH}")

        if conn is None: # Should not happen if previous steps are successful
            raise Exception("Failed to establish database connection.")

        cursor = conn.cursor()

        # Create tables based on schema definitions
        for table_name, schema in TABLE_SCHEMAS.items():
            try:
                cursor.execute(schema['create_sql'])
                logging.info(f"Table '{table_name}' created successfully." if cursor.rowcount == -1 else f"Table '{table_name}' already exists. Skipping creation.")
            except sqlite3.OperationalError as e:
                logging.error(f"Error creating table {table_name}: {e}")
                # This might happen if a column definition is wrong. Log and continue for other tables.
            except Exception as e:
                logging.critical(f"CRITICAL ERROR: Failed to create table {table_name}: {e}")
                raise # Stop if a critical table cannot be created

        # Insert predefined categories if the Categories table is empty
        try:
            cursor.execute("SELECT COUNT(*) FROM Categories")
            count = cursor.fetchone()[0]
            if count == 0:
                logging.info("Categories table is empty. Inserting predefined categories.")
                for category in PREDEFINED_CATEGORIES:
                    cursor.execute("INSERT INTO Categories (CategoryName) VALUES (?)", (category['CategoryName'],))
                conn.commit()
                logging.info("Predefined categories inserted successfully.")
            else:
                logging.info("Categories table already contains data. Skipping predefined category insertion.")
        except sqlite3.OperationalError as e:
            logging.warning(f"Could not check/insert predefined categories (Categories table might not exist or be corrupt): {e}")
        except Exception as e:
            logging.error(f"Error inserting predefined categories: {e}")

        conn.commit()
        logging.info("Database initialization process completed.")

    except sqlite3.Error as e:
        logging.critical(f"Overall Database initialization failed: {e}")
        raise # Re-raise the exception to be caught by the orchestrator/GUI
    except Exception as e:
        logging.critical(f"An unexpected error occurred during database initialization: {e}")
        raise # Re-raise the exception
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    try:
        initialize_database()
        print("Database initialization script finished. Check DebugLog.txt for details.")
    except Exception as e:
        print(f"Database initialization failed: {e}. See DebugLog.txt for critical errors.")

