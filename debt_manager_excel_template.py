# debt_manager_excel_template.py
# Purpose: Creates or updates the Excel dashboard template (DebtDashboard.xlsx)
#          with necessary sheets and headers, and adds basic visual elements.
# Deploy in: C:\DebtTracker
# Version: 2.4 (2025-07-19) - Fixed 'AttributeError: 'list' object has no attribute 'add''
#          by correctly clearing the worksheet's internal table collection.
#          Ensures only openpyxl is used for all Excel operations, including button styling.

import os
import logging
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles.colors import Color
# No xlwings import here!

from config import EXCEL_PATH, TABLE_SCHEMAS, LOG_FILE, LOG_DIR

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

def create_excel_template():
    """
    Creates or updates the DebtDashboard.xlsx Excel file.
    Uses openpyxl for all operations, including creating sheets, headers,
    and styling cells to visually represent dashboard navigation buttons.
    """
    logging.info("Starting Excel template creation/update process.")

    workbook_exists = os.path.exists(EXCEL_PATH)
    wb = None # Use 'wb' consistently for openpyxl workbook

    try:
        # --- Part 1: Create/Update Sheets and Headers using openpyxl ---
        if workbook_exists:
            try:
                wb = load_workbook(EXCEL_PATH)
                logging.info(f"Opened existing workbook with openpyxl: {EXCEL_PATH}")
            except Exception as e:
                logging.warning(f"Could not load existing workbook with openpyxl: {e}. Creating new one.")
                wb = Workbook()
                workbook_exists = False # Treat as new if loading fails
        else:
            wb = Workbook()
            # If new workbook, ensure a default sheet is present or remove it
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
            logging.info(f"Created new Excel workbook with openpyxl: {EXCEL_PATH}")

        # Ensure a 'Dashboard' sheet exists and is the first one
        if 'Dashboard' not in wb.sheetnames:
            dashboard_sheet = wb.create_sheet("Dashboard", 0)
            logging.info("Created 'Dashboard' sheet.")
        else:
            dashboard_sheet = wb['Dashboard']
            wb.move_sheet(dashboard_sheet, offset=-wb.sheetnames.index('Dashboard')) # Move to first position
            logging.info("Ensured 'Dashboard' sheet is first.")

        # Clear existing content on Dashboard sheet for fresh start
        # First, unmerge all merged cells to avoid 'MergedCell' error
        for merged_range in list(dashboard_sheet.merged_cells): # Iterate over a copy of merged ranges
            dashboard_sheet.unmerge_cells(str(merged_range))

        # Clear cell values and styles in a reasonable range
        for row_idx in range(1, 100): # Clear up to row 99
            for col_idx in range(1, 20): # Clear up to column T
                cell = dashboard_sheet.cell(row=row_idx, column=col_idx)
                cell.value = None
                cell.style = 'Normal' # Reset style
                cell.font = Font()
                cell.fill = PatternFill()
                cell.border = Border()
                cell.alignment = Alignment()

        logging.info("Prepared workbook: renamed first sheet to 'Dashboard' and cleared its content.")

        # Create/update data sheets and add headers
        for table_name, schema in TABLE_SCHEMAS.items():
            sheet_name = table_name # Use table name as sheet name
            if sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(sheet_name)
                logging.warning(f"Sheet '{sheet_name}' not found in Excel, creating it.")
            else:
                ws = wb[sheet_name]

            # Clear existing data (keep headers)
            for row_idx in range(2, ws.max_row + 1):
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).value = None

            # Add headers if not present or ensure they are correct
            headers = schema['excel_columns']
            for col_idx, header_text in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header_text)
                # Apply basic styling to headers
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Auto-adjust column widths based on header length
            for col_idx, header_text in enumerate(headers, 1):
                max_length = len(header_text) # Start with header length
                adjusted_width = (max_length + 2) if max_length > 0 else 15
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

            # --- CRITICAL FIX: Clear all existing tables from the worksheet's internal _tables collection ---
            # This prevents the 'AttributeError: 'list' object has no attribute 'add''
            # by ensuring _tables is a proper openpyxl collection before adding a new Table object.
            # We iterate and remove existing tables.
            if hasattr(ws, '_tables'):
                # Create a list of table names to remove to avoid modifying while iterating
                tables_to_remove = [table.name for table in ws._tables]
                for table_name_to_remove in tables_to_remove:
                    try:
                        # openpyxl's remove_table method expects the table name
                        ws.remove_table(table_name_to_remove)
                        logging.info(f"Removed old table '{table_name_to_remove}' from sheet '{sheet_name}'.")
                    except Exception as e:
                        logging.warning(f"Could not remove table '{table_name_to_remove}' from sheet '{sheet_name}': {e}")

            # Now, add the new Table object
            table_display_name = f"{sheet_name}Table"
            try:
                # The table reference should only include the header row at this stage of template creation
                tab_ref = f"A1:{get_column_letter(len(headers))}1"
                table = Table(displayName=table_display_name, ref=tab_ref)
                style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                table.tableStyleInfo = style
                ws.add_table(table)
                logging.info(f"Added new table '{table_display_name}' to sheet '{sheet_name}'.")
            except Exception as e:
                logging.error(f"Error adding table '{table_display_name}' to sheet {sheet_name}: {e}", exc_info=True)
                raise # Re-raise to ensure the orchestrator knows this step failed.


        # --- Part 2: Add Visual "Buttons" to Dashboard using openpyxl cell styling ---
        # Define common styles for "buttons"
        button_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Light blue
        button_font = Font(bold=True, size=10, color="000000") # Black text
        button_alignment = Alignment(horizontal='center', vertical='center')
        button_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                               top=Side(style='thin'), bottom=Side(style='thin'))

        # Helper function to create a visual "button" by styling cells
        def create_visual_button(sheet, start_row, start_col, text, width_cells=3, height_cells=2):
            # Merge cells to create a larger button area
            end_row = start_row + height_cells - 1
            end_col = start_col + width_cells - 1
            sheet.merge_cells(start_row=start_row, start_column=start_col,
                              end_row=end_row, end_column=end_col)

            cell = sheet.cell(row=start_row, column=start_col)
            cell.value = text
            cell.font = button_font
            cell.fill = button_fill
            cell.alignment = button_alignment

            # Apply border to the merged cell area
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    sheet.cell(row=r, column=c).border = button_border

            # Set row height and column width for the button area
            sheet.row_dimensions[start_row].height = 30 # Adjust height as needed
            for col_idx in range(start_col, end_col + 1):
                sheet.column_dimensions[get_column_letter(col_idx)].width = 15 # Adjust width as needed

        # Calculate positions for buttons
        start_row = 3
        start_col = 2 # Column B
        col_offset = 4 # Number of cells between buttons horizontally
        row_offset = 3 # Number of rows between button rows vertically

        # Row 1 of buttons
        create_visual_button(dashboard_sheet, start_row, start_col, "Dashboard")
        create_visual_button(dashboard_sheet, start_row, start_col + col_offset, "Debts Tab")
        create_visual_button(dashboard_sheet, start_row, start_col + 2 * col_offset, "Accounts Tab")

        # Row 2 of buttons
        create_visual_button(dashboard_sheet, start_row + row_offset, start_col, "Payments Tab")
        create_visual_button(dashboard_sheet, start_row + row_offset, start_col + col_offset, "Goals Tab")
        create_visual_button(dashboard_sheet, start_row + row_offset, start_col + 2 * col_offset, "Assets Tab")

        # Row 3 of buttons
        create_visual_button(dashboard_sheet, start_row + 2 * row_offset, start_col, "Revenue Tab")
        create_visual_button(dashboard_sheet, start_row + 2 * row_offset, start_col + col_offset, "Categories Tab")
        create_visual_button(dashboard_sheet, start_row + 2 * row_offset, start_col + 2 * col_offset, "Reports Tab")

        # Row 4 of derived debt tabs
        create_visual_button(dashboard_sheet, start_row + 3 * row_offset, start_col, "Bills Tab")
        create_visual_button(dashboard_sheet, start_row + 3 * row_offset, start_col + col_offset, "Credit Cards Tab")
        create_visual_button(dashboard_sheet, start_row + 3 * row_offset, start_col + 2 * col_offset, "Loans Tab")

        # Row 5 of derived debt tabs
        create_visual_button(dashboard_sheet, start_row + 4 * row_offset, start_col, "Collections Tab")

        # Save the workbook
        wb.save(EXCEL_PATH)
        logging.info(f"Excel template (sheets, headers, and visual buttons) saved with openpyxl to: {EXCEL_PATH}")

    except Exception as e:
        logging.error(f"Error during Excel template creation/drawing: {e}", exc_info=True) # Log full traceback
        raise # Re-raise to be caught by orchestrator
    finally:
        # No xlwings or COM objects to clean up here, just ensure openpyxl workbook is handled.
        pass

    logging.info("Excel template creation/update process completed.")

if __name__ == "__main__":
    try:
        create_excel_template()
        print(f"Excel template created/updated at: {EXCEL_PATH}. Check DebugLog.txt for details.")
    except Exception as e:
        print(f"Failed to create/update Excel template: {e}. See DebugLog.txt for errors.")
