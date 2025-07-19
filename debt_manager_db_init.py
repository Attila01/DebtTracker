# debt_manager_db_init.py
# Purpose: Handles the initial creation of the SQLite database and its tables,
#          including inserting predefined categories.
# Deploy in: C:\DebtTracker
# Version: 1.5 (2025-07-19) - Confirmed dynamic schema generation handles new columns from config.py.
#          Re-engineered table creation to dynamically build SQL
#          from TABLE_SCHEMAS['columns'] instead of using a 'create_sql' key.
#          Incorporated logic to add missing columns to existing tables.
#          Improved logging for database initialization.

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
    - Adds any missing columns to existing tables.
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
                conn.row_factory = sqlite3.Row # Ensure row_factory is set for consistency
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
            conn.row_factory = sqlite3.Row # Ensure row_factory is set for consistency
            logging.info(f"Database file created successfully: {DB_PATH}")

        if conn is None: # Should not happen if previous steps are successful
            raise Exception("Failed to establish database connection.")

        cursor = conn.cursor()

        # Create tables and add missing columns based on schema definitions
        for table_name, schema in TABLE_SCHEMAS.items():
            columns_sql = []
            for col in schema['columns']:
                col_def = f"{col['name']} {col['type']}"
                if col.get('primary_key'):
                    col_def += ' PRIMARY KEY'
                    if col.get('autoincrement'):
                        col_def += ' AUTOINCREMENT'
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

            create_table_sql = f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(columns_sql)});"

            try:
                cursor.execute(create_table_sql)
                conn.commit()
                logging.info(f"Table '{table_name}' ensured to exist (created if not present).")

                # After creating/ensuring table, check for missing columns and add them
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
                            logging.info(f"Added missing column to {table_name}: {col_def}")
                            conn.commit()
                        except sqlite3.Error as e:
                            logging.warning(f"Could not add column {col['name']} to {table_name}: {e}")

            except sqlite3.OperationalError as e:
                logging.error(f"Error creating/updating table {table_name}: {e}")
                raise # Re-raise if table creation itself fails
            except Exception as e:
                logging.critical(f"CRITICAL ERROR: Failed to create/update table {table_name}: {e}")
                raise # Stop if a critical table cannot be created

        # Special handling for Categories table: insert predefined categories if empty
        # This block should be outside the main table creation loop to ensure Categories table is ready
        cursor.execute("SELECT COUNT(*) FROM Categories")
        if cursor.fetchone()[0] == 0:
            logging.info("Categories table is empty. Inserting predefined categories.")
            for category_name in PREDEFINED_CATEGORIES: # PREDEFINED_CATEGORIES is a list of strings
                try:
                    cursor.execute("INSERT INTO Categories (CategoryName) VALUES (?)", (category_name,))
                    conn.commit()
                except sqlite3.IntegrityError: # In case of unique constraint violation
                    logging.warning(f"Category '{category_name}' already exists, skipping insertion.")
            logging.info("Predefined categories inserted successfully.")
        else:
            logging.info("Categories table already contains data. Skipping predefined category insertion.")

        conn.commit() # Final commit for any pending changes
        logging.info("Database initialization process completed.")

    except sqlite3.Error as e:
        logging.critical(f"Overall Database initialization failed: {e}", exc_info=True)
        raise # Re-raise the exception to be caught by the orchestrator/GUI
    except Exception as e:
        logging.critical(f"An unexpected error occurred during database initialization: {e}", exc_info=True)
        raise # Re-raise the exception
    finally:
        if conn:
            conn.close()
            logging.info("Database connection closed after initialization.")

if __name__ == "__main__":
    try:
        initialize_database()
        print("Database initialization script finished. Check DebugLog.txt for details.")
    except Exception as e:
        print(f"Database initialization failed: {e}. See DebugLog.txt for critical errors.")
