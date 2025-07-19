# debt_manager_excel_sync.py
# Purpose: Synchronizes data between the SQLite database (debt_manager.db)
#          and the Excel dashboard (DebtDashboard.xlsx).
# Deploy in: C:\DebtTracker
# Version: 1.3 (2025-07-19) - Enhanced data sanitization to explicitly handle null bytes and
#          other problematic characters, resolving IllegalCharacterError.
#          Confirmed compatibility with updated config.py schemas
#          and db_manager queries. Uses openpyxl for Excel and
#          sqlite3 for database operations. Designed to replace PowerShell sync
#          logic and work seamlessly with the openpyxl-based Excel template.

import os
import logging
import sqlite3
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import pandas as pd
import re # Import regex for sanitization

from config import DB_PATH, EXCEL_PATH, TABLE_SCHEMAS, LOG_FILE, LOG_DIR
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

def sanitize_excel_string(value):
    """
    Removes characters that are illegal or problematic in Excel cell values,
    especially null bytes and other control characters.
    Ensures the value is a string before sanitizing.
    """
    if value is None:
        return "" # Convert None to empty string for Excel
    if not isinstance(value, str):
        return value # Return non-string values as is (numbers, booleans, dates handled by openpyxl)

    # Explicitly replace null bytes, which are a common cause of IllegalCharacterError
    cleaned_text = value.replace('\x00', '')

    # Remove other control characters (excluding common ones like tab, newline, carriage return)
    # This regex matches characters in the ASCII range 0-31, excluding 9 (tab), 10 (LF), 13 (CR)
    cleaned_text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', cleaned_text)

    # Optionally, remove characters that might cause issues in some Excel versions
    # This is a broader sweep for potentially problematic Unicode characters
    # cleaned_text = ''.join(char for char in cleaned_text if char.isprintable() or char in ('\n', '\t', '\r'))

    return cleaned_text

def sqlite_to_excel():
    """
    Synchronizes data from SQLite database tables to corresponding Excel sheets.
    This function overwrites data in Excel sheets with the latest data from SQLite.
    """
    logging.info("Starting sqlite_to_excel sync...")
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # Load the Excel workbook
        if os.path.exists(EXCEL_PATH):
            try:
                wb = load_workbook(EXCEL_PATH)
            except Exception as e:
                logging.warning(f"Could not load existing Excel workbook with openpyxl: {e}. Creating new one.")
                wb = Workbook()
        else:
            wb = Workbook()

        # Remove default 'Sheet' if it exists in a new workbook
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

        for table_name, schema in TABLE_SCHEMAS.items():
            sheet_name = table_name # Use table name as sheet name

            # Ensure the sheet exists in the workbook
            if sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(sheet_name)
                logging.warning(f"Sheet '{sheet_name}' not found in Excel, creating it for sync.")
            else:
                ws = wb[sheet_name]

            # Fetch data from SQLite using db_manager's get_table_data for consistency
            df_sqlite = db_manager.get_table_data(table_name)

            # Get column names from the DataFrame, which reflects the query output (including joined cols)
            column_names = list(df_sqlite.columns)

            # Clear existing data in the Excel sheet (keep headers if they exist)
            for row_idx in range(ws.max_row, 1, -1):
                ws.delete_rows(row_idx)

            # Write headers to the first row if not present or if they need updating
            excel_headers = schema['excel_columns']
            for col_idx, header in enumerate(excel_headers, 1):
                ws.cell(row=1, column=col_idx, value=header)

            # Write data rows
            for index, row_data in df_sqlite.iterrows():
                mapped_row = []
                for col_name in excel_headers:
                    if col_name in row_data:
                        # Sanitize the value before appending to Excel
                        sanitized_value = sanitize_excel_string(row_data[col_name])
                        mapped_row.append(sanitized_value)
                    else:
                        mapped_row.append(None) # Append None if column not in DataFrame
                ws.append(mapped_row)

            # Auto-adjust column widths based on content
            for col_idx, header_text in enumerate(excel_headers, 1):
                max_length = len(header_text)
                for row_data_cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row_data_cell:
                        if cell.value is not None:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                adjusted_width = (max_length + 2) if max_length > 0 else 15
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

            logging.info(f"Synced data from SQLite table '{table_name}' to Excel sheet '{sheet_name}'.")

        wb.save(EXCEL_PATH)
        logging.info("sqlite_to_excel sync completed successfully.")

    except Exception as e:
        logging.error(f"Error during sqlite_to_excel sync: {e}", exc_info=True)
        raise # Re-raise to be caught by orchestrator
    finally:
        if conn:
            conn.close()
            logging.info("SQLite connection closed after sync.")

def excel_to_sqlite():
    """
    Synchronizes data from Excel sheets to corresponding SQLite database tables.
    This function overwrites data in SQLite tables with the latest data from Excel.
    """
    logging.info("Starting excel_to_sqlite sync...")
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        if not os.path.exists(EXCEL_PATH):
            logging.warning(f"Excel file not found at {EXCEL_PATH}. Skipping excel_to_sqlite sync.")
            return

        wb = load_workbook(EXCEL_PATH)

        for table_name, schema in TABLE_SCHEMAS.items():
            sheet_name = table_name
            if sheet_name not in wb.sheetnames:
                logging.warning(f"Excel sheet '{sheet_name}' not found. Skipping sync for this table.")
                continue

            ws = wb[sheet_name]

            # Read data from Excel, assuming first row is headers
            data = ws.values
            columns = next(data) # Get headers from the first row
            df = pd.DataFrame(data, columns=columns)

            # Ensure column names match SQLite table schema (case-insensitive if needed, but strict here)
            # Filter DataFrame to only include columns relevant to the SQLite table
            sqlite_columns_expected = [col['name'] for col in schema['columns']]

            # Delete existing data in SQLite table
            cursor.execute(f"DELETE FROM {table_name}")

            # Insert data from DataFrame into SQLite
            # Filter df to only include columns that exist in the SQLite table schema
            df_filtered = df[[col for col in sqlite_columns_expected if col in df.columns]]

            # Prepare for insertion: create placeholders and column names string
            cols = ", ".join(df_filtered.columns)
            placeholders = ", ".join(["?" for _ in df_filtered.columns])

            # Convert DataFrame rows to a list of tuples for executemany
            data_to_insert = [tuple(row) for row in df_filtered.values]

            if data_to_insert: # Only execute if there's data to insert
                cursor.executemany(f"INSERT INTO {table_name} ({cols}) VALUES ({placeholders})", data_to_insert)

            conn.commit()
            logging.info(f"Synced data from Excel sheet '{sheet_name}' to SQLite table '{table_name}'.")

    except Exception as e:
        logging.error(f"Error during excel_to_sqlite sync: {e}", exc_info=True)
        raise # Re-raise to be caught by orchestrator
    finally:
        if conn:
            conn.close()
            logging.info("SQLite connection closed after sync.")

if __name__ == "__main__":
    try:
        sqlite_to_excel()
        print("Excel-SQLite synchronization script executed. Check DebugLog.txt for details.")
    except Exception as e:
        print(f"Failed to execute Excel-SQLite synchronization: {e}. See DebugLog.txt for errors.")
