# debt_manager_excel_template.py
# Purpose: Creates or updates the Debt Management System Excel Dashboard (DebtDashboard.xlsx).
#          This is a Python alternative to CreateExcelTemplate.ps1, now using xlwings for drawing.
# Deploy in: C:\DebtTracker
# Version: 1.4 (2025-07-18) - Corrected xlwings TextFrame2 alignment properties.
#                             Uses TextFrame2.TextRange.ParagraphFormat.Alignment and TextFrame2.VerticalAnchor
#                             with appropriate Mso constants.
#                             Requires Microsoft Excel to be installed.

import os
import logging
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import xlwings as xw # Import xlwings for Excel automation

# Import configuration from config.py
from config import EXCEL_PATH, LOG_FILE, LOG_DIR, TABLE_SCHEMAS

# Ensure log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s: %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE, mode='a'),
                        logging.StreamHandler()
                    ])

def create_excel_template():
    """
    Creates or updates the Debt Management System Excel Dashboard template.
    This function will:
    1. Create a new workbook or load an existing one using openpyxl.
    2. Set up the 'Dashboard' sheet with summary sections.
    3. Use xlwings to draw the data flow diagram on the Dashboard sheet.
    4. Create/update individual data sheets with appropriate headers using openpyxl.
    """
    logging.info("Starting Excel template creation/update process.")

    app = None
    wb = None
    try:
        # --- Part 1: Initial Workbook Setup with openpyxl (for basic sheet creation/headers) ---
        # This part still uses openpyxl as it's efficient for sheet content and styling.
        # We will then use xlwings for the drawing objects.
        workbook = None
        if os.path.exists(EXCEL_PATH):
            try:
                workbook = load_workbook(EXCEL_PATH)
                logging.info(f"Opened existing workbook with openpyxl: {EXCEL_PATH}")
            except Exception as e:
                logging.warning(f"Error opening existing workbook with openpyxl: {e}. Creating a new one.")
                workbook = Workbook()
        else:
            workbook = Workbook()
            logging.info(f"Created new workbook with openpyxl: {EXCEL_PATH}")

        # Ensure only one sheet initially, named "Dashboard"
        while len(workbook.sheetnames) > 1:
            workbook.remove(workbook[workbook.sheetnames[1]])

        dashboard_sheet = workbook.active
        dashboard_sheet.title = "Dashboard"
        logging.info("Prepared workbook: renamed first sheet to 'Dashboard'.")

        # Configure Dashboard Sheet content (titles, headers, placeholders)
        dashboard_sheet.sheet_view.showGridLines = False # Hide gridlines for a cleaner look

        # Set default font for the entire sheet
        for row in dashboard_sheet.iter_rows():
            for cell in row:
                cell.font = Font(name='Inter', size=10)

        # Title
        dashboard_sheet['A1'] = "Financial Dashboard"
        dashboard_sheet['A1'].font = Font(name='Inter', size=24, bold=True)
        dashboard_sheet['A1'].fill = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid") # Light blue
        dashboard_sheet.merge_cells('A1:F1')
        dashboard_sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')

        # Helper for section styling (openpyxl)
        def set_section_style_openpyxl(sheet, start_row, start_col, title, headers):
            """Applies styling and adds title/headers for a dashboard section."""
            title_cell = sheet.cell(row=start_row, column=start_col)
            title_cell.value = title
            title_cell.font = Font(name='Inter', size=14, bold=True)
            title_cell.fill = PatternFill(start_color="C8DCEF", end_color="C8DCEF", fill_type="solid") # Slightly darker blue

            # Merge cells for the title
            end_col_letter = get_column_letter(start_col + len(headers) - 1)
            sheet.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + len(headers) - 1)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')

            # Add headers for the section
            header_row = start_row + 1
            for i, header_text in enumerate(headers):
                header_cell = sheet.cell(row=header_row, column=start_col + i)
                header_cell.value = header_text
                header_cell.font = Font(name='Inter', bold=True)
                header_cell.fill = PatternFill(start_color="DCEDFA", end_color="DCEDFA", fill_type="solid") # Lighter blue
                header_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                header_cell.alignment = Alignment(horizontal='center', vertical='center')

            # Set column widths for headers
            for i, header_text in enumerate(headers):
                col_letter = get_column_letter(start_col + i)
                sheet.column_dimensions[col_letter].width = max(len(header_text) + 2, 15) # Ensure minimum width

            return header_row + 1 # Return next available row after headers

        # Dashboard Sections
        current_row = 3

        # Debt Overview
        debt_overview_headers = ["Metric", "Value"]
        current_row = set_section_style_openpyxl(dashboard_sheet, current_row, 1, "Debt Overview", debt_overview_headers)
        dashboard_sheet.cell(row=current_row, column=1, value="Total Outstanding Debt:")
        dashboard_sheet.cell(row=current_row + 1, column=1, value="Original Total Debt:")
        dashboard_sheet.cell(row=current_row + 2, column=1, value="Number of Active Debts:")
        dashboard_sheet.cell(row=current_row + 3, column=1, value="Total Paid on Debts:")

        # Placeholders for values
        for i in range(4):
            dashboard_sheet.cell(row=current_row + i, column=2, value="$0.00").number_format = "$#,##0.00"
        dashboard_sheet.cell(row=current_row + 2, column=2, value="0").number_format = "0"

        current_row += 6 # Move to next section

        # Snowball Progress
        snowball_headers = TABLE_SCHEMAS['Debts']['excel_columns'] + ['Projected Payment']
        current_row = set_section_style_openpyxl(dashboard_sheet, current_row, 1, "Snowball Progress", snowball_headers)
        current_row += 2 # Move past headers

        # Cash Flow Summary
        cash_flow_headers = ["Month", "Income", "Expenses", "Net Flow"]
        current_row = set_section_style_openpyxl(dashboard_sheet, current_row, 1, "Cash Flow Summary", cash_flow_headers)
        current_row += 2

        # Account Balances
        account_balances_headers = TABLE_SCHEMAS['Accounts']['excel_columns']
        current_row = set_section_style_openpyxl(dashboard_sheet, current_row, 1, "Account Balances", account_balances_headers)
        current_row += 2

        # Goal Tracker
        goal_tracker_headers = TABLE_SCHEMAS['Goals']['excel_columns']
        current_row = set_section_style_openpyxl(dashboard_sheet, current_row, 1, "Goal Tracker", goal_tracker_headers)
        current_row += 2

        # Benchmark Milestones
        benchmark_headers = ["Milestone", "Target Date", "Achieved Date", "Status"]
        current_row = set_section_style_openpyxl(dashboard_sheet, current_row, 1, "Benchmark Milestones", benchmark_headers)
        current_row += 2

        # --- Create Data Sheets (using openpyxl) ---
        for table_name, schema_info in TABLE_SCHEMAS.items():
            if table_name == "Dashboard":
                continue

            if table_name in workbook.sheetnames:
                worksheet = workbook[table_name]
                worksheet.delete_rows(1, worksheet.max_row) # Clear existing content
                logging.info(f"Cleared existing sheet: {table_name}")
            else:
                worksheet = workbook.create_sheet(title=table_name)
                logging.info(f"Created sheet: {table_name}")

            headers = schema_info['excel_columns']
            if table_name == 'Debts':
                headers = headers + ['Projected Payment']

            for i, header_text in enumerate(headers):
                cell = worksheet.cell(row=1, column=i + 1)
                cell.value = header_text
                cell.font = Font(name='Inter', bold=True)
                cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Light Blue
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')

            worksheet.row_dimensions[1].height = 20

            for col in range(1, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(col)].width = max(len(headers[col-1]) + 2, 15)

        # Save the workbook with openpyxl before passing to xlwings
        workbook.save(EXCEL_PATH)
        logging.info(f"Excel workbook content (sheets, headers) saved with openpyxl to: {EXCEL_PATH}")

        # --- Part 2: Drawing Data Flow Diagram with xlwings ---
        # Launch Excel application (hidden)
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        xl_dashboard_sheet = wb.sheets["Dashboard"]
        logging.info("Opened workbook with xlwings for drawing.")

        # Define shape properties (in points, approximate)
        shape_width_pt = 110 # approx 1.5 inches
        shape_height_pt = 35 # approx 0.5 inches
        vertical_spacing_pt = 50
        horizontal_spacing_pt = 150

        # Calculate starting point for the diagram (relative to top-left of the sheet)
        # Using cell H3 as reference for diagram title
        # Column H starts roughly at 7 * default_col_width_points
        # Row 3 starts roughly at 2 * default_row_height_points
        # Let's place the diagram title at H3 directly using xlwings for consistency
        xl_dashboard_sheet.range('H3').value = "System Data Flow"
        xl_dashboard_sheet.range('H3').api.Font.Size = 16
        xl_dashboard_sheet.range('H3').api.Font.Bold = True
        xl_dashboard_sheet.range('H3').api.HorizontalAlignment = -4108 # xlCenter
        xl_dashboard_sheet.range('H3:P3').merge() # Merge cells for title

        # Starting position for shapes (approximate top-left corner of H5)
        start_x_pt = xl_dashboard_sheet.range('H5').left # Get left coordinate of H5
        start_y_pt = xl_dashboard_sheet.range('H5').top # Get top coordinate of H5

        # Define mso constants for clarity (these are COM enum values)
        msoShapeRoundedRectangle = 5
        msoConnectorStraight = 1
        msoConnectionSiteBottom = 3
        msoConnectionSiteTop = 1
        msoConnectionSiteRight = 2
        msoConnectionSiteLeft = 4
        msoArrowheadTriangle = 3
        xlCenter = -4108 # Horizontal and Vertical Alignment constant for VBA
        msoAlignCenter = 2 # MsoParagraphAlignment.msoAlignCenter
        msoAnchorMiddle = 3 # MsoVerticalAnchor.msoAnchorMiddle

        # Helper to add a rounded rectangle shape (using xlwings COM API)
        def add_rounded_rectangle_xl(sheet, x_pt, y_pt, width_pt, height_pt, text):
            # sheet.api.Shapes.AddShape(Type, Left, Top, Width, Height)
            shape = sheet.api.Shapes.AddShape(msoShapeRoundedRectangle, x_pt, y_pt, width_pt, height_pt)

            # Set text and alignment
            shape.TextFrame2.TextRange.Text = text
            # Corrected: Use ParagraphFormat.Alignment for horizontal text alignment
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            # Corrected: Use VerticalAnchor for vertical alignment of the text frame within the shape
            shape.TextFrame2.VerticalAnchor = msoAnchorMiddle

            # Set font properties via TextFrame2
            shape.TextFrame2.TextRange.Font.Name = 'Inter'
            shape.TextFrame2.TextRange.Font.Size = 9
            shape.TextFrame2.TextRange.Font.Bold = True

            # Set fill and line colors
            shape.Fill.ForeColor.RGB = 0xFFF8F0 # RGB for AliceBlue (reversed for VBA)
            shape.Line.Weight = 0.75 # in points
            shape.Line.ForeColor.RGB = 0x808080 # RGB for Grey (reversed for VBA)
            return shape

        # Helper to add an arrow (using xlwings COM API)
        def add_arrow_xl(sheet, start_shape, end_shape):
            # sheet.api.Shapes.AddConnector(Type, Left, Top, Width, Height) - initial dummy values
            connector = sheet.api.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
            connector.ConnectorFormat.BeginConnect(start_shape, msoConnectionSiteBottom)
            connector.ConnectorFormat.EndConnect(end_shape, msoConnectionSiteTop)
            connector.RerouteConnections() # Recalculate position based on connected shapes

            connector.Line.Weight = 1.5
            connector.Line.ForeColor.RGB = 0x8B0000 # RGB for DarkBlue (reversed for VBA)
            connector.Line.EndArrowheadStyle = msoArrowheadTriangle
            return connector

        # Helper to add a horizontal arrow (using xlwings COM API)
        def add_horizontal_arrow_xl(sheet, start_shape, end_shape):
            connector = sheet.api.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
            connector.ConnectorFormat.BeginConnect(start_shape, msoConnectionSiteRight)
            connector.ConnectorFormat.EndConnect(end_shape, msoConnectionSiteLeft)
            connector.RerouteConnections()

            connector.Line.Weight = 1.5
            connector.Line.ForeColor.RGB = 0x8B0000 # RGB for DarkBlue (reversed for VBA)
            connector.Line.EndArrowheadStyle = msoArrowheadTriangle
            return connector

        # Define positions for shapes (point coordinates)
        col1_x = start_x_pt
        col2_x = start_x_pt + horizontal_spacing_pt
        col3_x = start_x_pt + (2 * horizontal_spacing_pt)

        row1_y = start_y_pt
        row2_y = row1_y + vertical_spacing_pt
        row3_y = row2_y + vertical_spacing_pt
        row4_y = row3_y + vertical_spacing_pt
        row5_y = row4_y + vertical_spacing_pt
        row6_y = row5_y + vertical_spacing_pt
        row7_y = row6_y + vertical_spacing_pt

        # Create Shapes
        revenue_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col1_x, row1_y, shape_width_pt, shape_height_pt, "1. Revenue Tab (Income Source)")
        categories_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col1_x, row2_y, shape_width_pt, shape_height_pt, "2. Categories Tab (Defines IDs)")
        debts_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col2_x, row3_y, shape_width_pt, shape_height_pt, "3. Debts Tab (Linked to CategoryID, Projected Payments)")
        payments_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col2_x, row4_y, shape_width_pt, shape_height_pt, "4. Payments Tab (Confirms Transactions)")
        accounts_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col2_x, row5_y, shape_width_pt, shape_height_pt, "5. Accounts Tab (Receives Allocations, Sub-Rows)")

        bills_tab_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col3_x, row3_y, shape_width_pt, shape_height_pt, "6. Bills Tab (CategoryID=1)")
        credit_cards_tab_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col3_x, row4_y, shape_width_pt, shape_height_pt, "7. Credit Cards Tab (CategoryID=2)")
        loans_tab_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col3_x, row5_y, shape_width_pt, shape_height_pt, "8. Loans Tab (CategoryID=3)")
        collections_tab_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col3_x, row6_y, shape_width_pt, shape_height_pt, "9. Collections Tab (CategoryID=4)")

        dashboard_central_shape = add_rounded_rectangle_xl(xl_dashboard_sheet, col2_x, row7_y, shape_width_pt, shape_height_pt, "10. Dashboard Tab (Central View)")

        # Add Arrows to show data flow
        add_arrow_xl(xl_dashboard_sheet, revenue_shape, accounts_shape)
        add_arrow_xl(xl_dashboard_sheet, categories_shape, debts_shape)
        add_arrow_xl(xl_dashboard_sheet, debts_shape, payments_shape)
        add_arrow_xl(xl_dashboard_sheet, payments_shape, accounts_shape)

        # Payments -> Debts (feedback loop - need to adjust connection sites for upward arrow)
        connector_feedback = xl_dashboard_sheet.api.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
        connector_feedback.ConnectorFormat.BeginConnect(payments_shape, msoConnectionSiteTop)
        connector_feedback.ConnectorFormat.EndConnect(debts_shape, msoConnectionSiteBottom)
        connector_feedback.RerouteConnections()
        connector_feedback.Line.Weight = 1.5
        connector_feedback.Line.ForeColor.RGB = 0x8B0000 # RGB for DarkBlue (reversed for VBA)
        connector_feedback.Line.EndArrowheadStyle = msoArrowheadTriangle

        # Debts -> Category-based tabs (horizontal flow)
        add_horizontal_arrow_xl(xl_dashboard_sheet, debts_shape, bills_tab_shape)
        add_horizontal_arrow_xl(xl_dashboard_sheet, debts_shape, credit_cards_tab_shape)
        add_horizontal_arrow_xl(xl_dashboard_sheet, debts_shape, loans_tab_shape)
        add_horizontal_arrow_xl(xl_dashboard_sheet, debts_shape, collections_tab_shape)

        # All major tabs -> Dashboard (simplified central connection)
        for shape in [revenue_shape, categories_shape, debts_shape, payments_shape, accounts_shape,
                      bills_tab_shape, credit_cards_tab_shape, loans_tab_shape, collections_tab_shape]:
            add_arrow_xl(xl_dashboard_sheet, shape, dashboard_central_shape)

        # Save the workbook and quit Excel application
        wb.save()
        logging.info(f"Excel template drawings added with xlwings to: {EXCEL_PATH}")

    except Exception as e:
        logging.error(f"Error during Excel template creation/drawing: {e}", exc_info=True)
        raise # Re-raise the exception to be handled by the calling script
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()
            logging.info("xlwings Excel application instance quit.")

    logging.info("Excel template creation/update process completed.")

if __name__ == "__main__":
    try:
        create_excel_template()
        print("Excel dashboard template created/updated successfully!")
    except Exception as e:
        print(f"Failed to create/update Excel template: {e}")

