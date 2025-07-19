# debt_manager_excel_sync.py
# Purpose: Synchronizes data between the SQLite database (debt_manager.db)
#          and the Excel dashboard (DebtDashboard.xlsx).
# Deploy in: C:\DebtTracker
# Version: 1.0 (2025-07-19) - Initial version. Uses openpyxl for Excel and
#          sqlite3 for database operations. Designed to replace PowerShell sync
#          logic and work seamlessly with the openpyxl-based Excel template.

import os
import logging
import sqlite3
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import pandas as pd # Using pandas for easier data handling between DB and Excel
import time # For potential delays if needed for file access

from config import DB_PATH, EXCEL_PATH, TABLE_SCHEMAS, LOG_FILE, LOG_DIR

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
        # If the workbook doesn't exist or is corrupted, create a new one.
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

            # Fetch data from SQLite
            cursor.execute(f"SELECT * FROM {table_name}")
            rows = cursor.fetchall()

            # Get column names from SQLite table schema
            column_names = [description[0] for description in cursor.description]

            # Clear existing data in the Excel sheet (keep headers if they exist)
            # Find the last row with data (assuming headers are in row 1)
            # Iterate from row 2 downwards and clear content
            for row_idx in range(ws.max_row, 1, -1):
                ws.delete_rows(row_idx)

            # Write headers to the first row if not present or if they need updating
            excel_headers = schema['excel_columns']
            for col_idx, header in enumerate(excel_headers, 1):
                ws.cell(row=1, column=col_idx, value=header)

            # Write data rows
            for row_data in rows:
                # Map SQLite row data to Excel columns based on schema['excel_columns']
                # This ensures the order and inclusion of columns matches the Excel template
                mapped_row = []
                for col_name in excel_headers:
                    try:
                        # Find the index of the column name in SQLite's fetched column_names
                        col_index_in_sqlite = column_names.index(col_name)
                        mapped_row.append(row_data[col_index_in_sqlite])
                    except ValueError:
                        # If a column exists in excel_columns but not in SQLite, append None
                        mapped_row.append(None)
                ws.append(mapped_row)

            # Auto-adjust column widths based on content
            for col_idx, header_text in enumerate(excel_headers, 1):
                max_length = len(header_text)
                # Iterate through data rows to find max length
                for row_data in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row_data:
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
    # Example usage:
    # You can call these functions based on your orchestration needs.
    # For instance, if you want to push data from SQLite to Excel on startup:
    try:
        sqlite_to_excel()
        # If you want to also pull data from Excel to SQLite after some user interaction:
        # excel_to_sqlite()
        print("Excel-SQLite synchronization script executed. Check DebugLog.txt for details.")
    except Exception as e:
        print(f"Failed to execute Excel-SQLite synchronization: {e}. See DebugLog.txt for errors.")

