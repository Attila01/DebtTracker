# debt_manager_csv_sync.py
# Purpose: Synchronizes data between the SQLite database (debt_manager.db)
#          and individual CSV files (in csv_data directory).
# Deploy in: C:\DebtTracker
# Version: 1.5 (2025-07-19) - Refactored for CSV synchronization instead of Excel.
#          Uses pandas for CSV read/write operations.
#          Enhanced data sanitization to remove all non-printable
#          characters (except standard whitespace) to resolve persistent IllegalCharacterError.

import os
import logging
import sqlite3
import pandas as pd
import re # Import regex for sanitization

from config import DB_PATH, CSV_DIR, TABLE_SCHEMAS, LOG_FILE, LOG_DIR
import debt_manager_db_manager as db_manager

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

def sanitize_csv_string(value):
    """
    Removes characters that are problematic in CSV files or might cause issues
    when read by other applications. This is a general cleanup.
    Ensures the value is a string before sanitizing.
    """
    if value is None:
        return "" # Convert None to empty string
    if not isinstance(value, str):
        return value # Return non-string values as is

    original_value = value # Keep original for logging

    # Remove all control characters (ASCII 0-31), except for tab (\t), newline (\n), carriage return (\r)
    cleaned_text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', value)

    # Remove leading/trailing whitespace that might cause issues
    cleaned_text = cleaned_text.strip()

    if original_value != cleaned_text:
        logging.debug(f"Sanitized string: Original='{original_value[:50]}...', Cleaned='{cleaned_text[:50]}...'")

    return cleaned_text

def sqlite_to_csv():
    """
    Synchronizes data from SQLite database tables to corresponding CSV files.
    Each table is saved as a separate CSV file in the CSV_DIR.
    """
    logging.info("Starting sqlite_to_csv sync...")
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)

        # Ensure CSV directory exists
        os.makedirs(CSV_DIR, exist_ok=True)

        for table_name, schema in TABLE_SCHEMAS.items():
            csv_file_path = os.path.join(CSV_DIR, f"{table_name}.csv")

            # Fetch data from SQLite using db_manager's get_table_data for consistency
            df_sqlite = db_manager.get_table_data(table_name)

            # Ensure columns are in the order defined in 'csv_columns' and sanitize data
            if not df_sqlite.empty:
                # Select and reorder columns based on 'csv_columns'
                # Also, apply sanitization to all string columns
                df_to_save = pd.DataFrame()
                for col_name in schema['csv_columns']:
                    if col_name in df_sqlite.columns:
                        # Apply sanitization to string columns, convert others to string for CSV compatibility
                        # pandas to_csv handles most types, but explicit string conversion and sanitization is safer.
                        if df_sqlite[col_name].dtype == 'object': # Typically string columns
                            df_to_save[col_name] = df_sqlite[col_name].apply(sanitize_csv_string)
                        else:
                            df_to_save[col_name] = df_sqlite[col_name]
                    else:
                        df_to_save[col_name] = None # Add missing columns as None

                # Save to CSV
                df_to_save.to_csv(csv_file_path, index=False, encoding='utf-8')
                logging.info(f"Synced data from SQLite table '{table_name}' to CSV file '{csv_file_path}'.")
            else:
                # If DataFrame is empty, still create an empty CSV with headers
                empty_df = pd.DataFrame(columns=schema['csv_columns'])
                empty_df.to_csv(csv_file_path, index=False, encoding='utf-8')
                logging.info(f"No data for '{table_name}'. Created empty CSV file '{csv_file_path}' with headers.")

        logging.info("sqlite_to_csv sync completed successfully.")

    except Exception as e:
        logging.error(f"Error during sqlite_to_csv sync: {e}", exc_info=True)
        raise # Re-raise to be caught by orchestrator
    finally:
        if conn:
            conn.close()
            logging.info("SQLite connection closed after sync.")

def csv_to_sqlite():
    """
    Synchronizes data from CSV files to corresponding SQLite database tables.
    Each CSV file is read and its data overwrites the corresponding SQLite table.
    """
    logging.info("Starting csv_to_sqlite sync...")
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        if not os.path.exists(CSV_DIR):
            logging.warning(f"CSV directory not found at {CSV_DIR}. Skipping csv_to_sqlite sync.")
            return

        for table_name, schema in TABLE_SCHEMAS.items():
            csv_file_path = os.path.join(CSV_DIR, f"{table_name}.csv")

            if not os.path.exists(csv_file_path):
                logging.warning(f"CSV file '{csv_file_path}' not found. Skipping sync for this table.")
                continue

            # Read data from CSV
            # `keep_default_na=False` prevents pandas from converting empty strings to NaN.
            # `dtype=str` for all columns initially to avoid type inference issues with mixed types,
            # then convert explicitly.
            df_csv = pd.read_csv(csv_file_path, encoding='utf-8', keep_default_na=False)

            # Filter DataFrame to only include columns relevant to the SQLite table schema
            sqlite_columns_expected = [col['name'] for col in schema['columns']]

            # Ensure all expected columns are in the DataFrame, add as empty string if missing
            for col_def in schema['columns']:
                col_name = col_def['name']
                if col_name not in df_csv.columns:
                    df_csv[col_name] = None # Add missing columns as None

            # Reorder columns to match SQLite schema order for insertion
            df_filtered = df_csv[sqlite_columns_expected]

            # Type conversion based on SQLite schema before insertion
            for col_def in schema['columns']:
                col_name = col_def['name']
                db_type = col_def['type']
                if col_name in df_filtered.columns:
                    try:
                        if db_type == 'INTEGER':
                            # Convert to numeric, then to Int64 to handle NaNs (None for SQLite)
                            df_filtered[col_name] = pd.to_numeric(df_filtered[col_name], errors='coerce').astype('Int64')
                        elif db_type == 'REAL':
                            df_filtered[col_name] = pd.to_numeric(df_filtered[col_name], errors='coerce')
                        # TEXT and other types are fine as objects/strings
                    except Exception as e:
                        logging.warning(f"Error converting column '{col_name}' to type '{db_type}' for table '{table_name}': {e}. Data might be inserted as is.")


            # Delete existing data in SQLite table
            cursor.execute(f"DELETE FROM {table_name}")

            # Prepare for insertion: create placeholders and column names string
            cols = ", ".join(df_filtered.columns)
            placeholders = ", ".join(["?" for _ in df_filtered.columns])

            # Convert DataFrame rows to a list of tuples for executemany
            # Convert pandas.NA/NaN to None for SQLite compatibility
            data_to_insert = []
            for row in df_filtered.itertuples(index=False):
                converted_row = []
                for val in row:
                    if pd.isna(val): # Check for pandas NaN/NaT
                        converted_row.append(None)
                    else:
                        converted_row.append(val)
                data_to_insert.append(tuple(converted_row))

            if data_to_insert: # Only execute if there's data to insert
                cursor.executemany(f"INSERT INTO {table_name} ({cols}) VALUES ({placeholders})", data_to_insert)

            conn.commit()
            logging.info(f"Synced data from CSV file '{csv_file_path}' to SQLite table '{table_name}'.")

    except Exception as e:
        logging.error(f"Error during csv_to_sqlite sync: {e}", exc_info=True)
        raise # Re-raise to be caught by orchestrator
    finally:
        if conn:
            conn.close()
            logging.info("SQLite connection closed after sync.")

if __name__ == "__main__":
    try:
        sqlite_to_csv()
        # You can call csv_to_sqlite() here if you want to test round-trip sync
        # csv_to_sqlite()
        print("CSV synchronization script executed. Check DebugLog.txt for details.")
    except Exception as e:
        print(f"Failed to execute CSV synchronization: {e}. See DebugLog.txt for errors.")