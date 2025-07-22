# debt_manager_db_update_schema.py
# Purpose: Updates the SQLite database schema based on definitions in config.py.
#          This script can create new tables and add missing columns to existing tables.
# Deploy in: C:\DebtTracker
# Version: 1.1 (2025-07-19) - Fixed KeyError: 'db_columns' by referencing 'columns' from TABLE_SCHEMAS.

import sqlite3
import os
import logging
from config import DB_PATH, TABLE_SCHEMAS, LOG_FILE, LOG_DIR

# Ensure log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s: %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE, mode='a'),
                        logging.StreamHandler()
                    ])

def get_db_connection():
    """Establishes and returns a SQLite database connection."""
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row # Allows accessing columns by name
        return conn
    except sqlite3.Error as e:
        logging.error(f"Database connection error: {e}")
        raise # Re-raise to be caught by caller

def update_database_schema():
    """
    Updates the SQLite database schema:
    - Creates tables if they don't exist based on TABLE_SCHEMAS.
    - Adds missing columns to existing tables if schema has evolved.
    """
    logging.info("Starting database schema update process.")

    # Ensure the database directory exists
    db_dir = os.path.dirname(DB_PATH)
    if not os.path.exists(db_dir):
        os.makedirs(db_dir)
        logging.info(f"Created database directory: {db_dir}")

    # Ensure the database file exists, if not, initialize it (creates empty tables)
    if not os.path.exists(DB_PATH):
        logging.warning(f"Database file not found at {DB_PATH}. Attempting to create it.")
        # This will create an empty database file, then tables will be added below.
        conn_temp = None
        try:
            conn_temp = sqlite3.connect(DB_PATH)
            conn_temp.close()
            logging.info(f"Empty database file created at {DB_PATH}.")
        except sqlite3.Error as e:
            logging.critical(f"Failed to create empty database file: {e}", exc_info=True)
            raise

    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # Iterate through all defined tables and create them if missing, or update if existing
        for table_name, schema_info in TABLE_SCHEMAS.items():
            # Correctly reference the 'columns' key from schema_info
            schema_columns_defs = schema_info['columns']
            primary_key_name = schema_info['primary_key'] # Get primary key name from schema

            # Check if table exists
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
            table_exists = cursor.fetchone()

            if not table_exists:
                # Construct CREATE TABLE statement
                columns_ddl = []
                for col_def in schema_columns_defs:
                    col_name = col_def['name']
                    col_type = col_def['type']
                    col_sql = f"{col_name} {col_type}"

                    if col_def.get('primary_key'):
                        col_sql += " PRIMARY KEY"
                        if col_def.get('autoincrement'):
                            col_sql += " AUTOINCREMENT"
                    if not col_def.get('nullable') and not col_def.get('primary_key'):
                        col_sql += " NOT NULL"
                    if 'default' in col_def:
                        default_val = col_def['default']
                        if isinstance(default_val, str):
                            col_sql += f" DEFAULT '{default_val}'"
                        else:
                            col_sql += f" DEFAULT {default_val}"
                    if col_def.get('unique'):
                        col_sql += " UNIQUE"
                    columns_ddl.append(col_sql)

                create_table_sql = f"CREATE TABLE {table_name} ({', '.join(columns_ddl)})"
                try:
                    cursor.execute(create_table_sql)
                    conn.commit()
                    logging.info(f"Table '{table_name}' created successfully.")
                except sqlite3.Error as e:
                    logging.error(f"Error creating table '{table_name}': {e}")
            else:
                logging.info(f"Table '{table_name}' already exists. Checking for missing columns.")
                # Check for missing columns and add them
                cursor.execute(f"PRAGMA table_info({table_name})")
                existing_columns_info = cursor.fetchall()
                existing_column_names = {col_info[1] for col_info in existing_columns_info}

                for col_def in schema_columns_defs:
                    col_name = col_def['name']
                    if col_name not in existing_column_names:
                        col_type = col_def['type']
                        add_column_sql = f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type}"

                        if not col_def.get('nullable') and not col_def.get('primary_key'):
                            add_column_sql += " NOT NULL"
                        if 'default' in col_def:
                            default_val = col_def['default']
                            if isinstance(default_val, str):
                                add_column_sql += f" DEFAULT '{default_val}'"
                            else:
                                add_column_sql += f" DEFAULT {default_val}"
                        if col_def.get('unique'):
                            add_column_sql += " UNIQUE"

                        try:
                            cursor.execute(add_column_sql)
                            conn.commit()
                            logging.info(f"Added column '{col_name}' to table '{table_name}'.")
                        except sqlite3.Error as e:
                            logging.error(f"Error adding column '{col_name}' to table '{table_name}': {e}")
                            # Continue even if one column fails, to try others

        logging.info("Database schema update process completed.")

    except sqlite3.Error as e:
        logging.critical(f"Critical database schema update error: {e}", exc_info=True)
        raise # Re-raise the exception to be handled by the calling script
    except Exception as e:
        logging.critical(f"An unexpected error occurred during database schema update: {e}", exc_info=True)
        raise
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    # This block allows the script to be run independently for testing
    try:
        update_database_schema()
        print("Database schema updated successfully.")
    except Exception as e:
        print(f"Database schema update failed: {e}")