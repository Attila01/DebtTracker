# debt_manager_orchestrator.py
# Purpose: Orchestrates the startup process for the Debt Management System.
#          It ensures the database schema is up-to-date, the Excel template is created/updated,
#          and then launches the main GUI application.
# Deploy in: C:\DebtTracker
# Version: 1.2 (2025-07-18) - Updated to accommodate xlwings usage in Excel template creation.

import os
import subprocess
import logging
import sys

# Import functions from other modules
# Ensure these modules are in the same directory or accessible via PYTHONPATH
try:
    from config import LOG_FILE, LOG_DIR, DB_PATH, EXCEL_PATH
    from debt_manager_db_update_schema import update_database_schema
    from debt_manager_excel_template import create_excel_template # Now importing Python version
    # No direct import for debt_manager_gui, as we'll run it as a separate process
except ImportError as e:
    print(f"Error importing necessary modules: {e}")
    print("Please ensure 'config.py', 'debt_manager_db_update_schema.py', and 'debt_manager_excel_template.py' are in the same directory.")
    print("Also, ensure 'xlwings' is installed if using Excel drawing features (`pip install xlwings`).")
    sys.exit(1)

# Ensure log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s: %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE, mode='a'),
                        logging.StreamHandler() # Also log to console
                    ])

def run_orchestrator():
    """
    Executes the full orchestration sequence:
    1. Updates the database schema.
    2. Creates/updates the Excel dashboard template.
    3. Launches the main GUI application.
    """
    logging.info("--- Starting Debt Management System Orchestrator ---")

    # Step 1: Update Database Schema
    try:
        logging.info("Step 1: Updating database schema...")
        update_database_schema()
        logging.info("Step 1: Database schema updated successfully.")
    except Exception as e:
        logging.critical(f"Step 1: Database schema update failed: {e}", exc_info=True)
        print(f"ERROR: Database schema update failed. Check DebugLog.txt for details. Error: {e}")
        sys.exit(1) # Exit if database update fails

    # Step 2: Create/Update Excel Template (using Python script)
    try:
        logging.info("Step 2: Creating/updating Excel template...")
        # create_excel_template now handles its own Excel application instance via xlwings.
        # It will also attempt to close any open Excel instances it starts.
        create_excel_template() # Call the Python function directly
        logging.info("Step 2: Excel template created/updated successfully.")
    except Exception as e:
        logging.critical(f"Step 2: Excel template creation/update failed: {e}", exc_info=True)
        print(f"ERROR: Excel template creation/update failed. Ensure Microsoft Excel is installed and check DebugLog.txt. Error: {e}")
        sys.exit(1) # Exit if Excel template creation fails

    # Step 3: Launch GUI Application
    try:
        logging.info("Step 3: Launching Debt Management System GUI...")
        # Use subprocess.run to execute the GUI script as a separate process.
        # This allows the orchestrator to finish its tasks and the GUI to run independently.
        # We use sys.executable to ensure the correct Python interpreter is used.
        gui_script_path = os.path.join(os.path.dirname(__file__), 'debt_manager_gui.py')

        # Check if the GUI script exists
        if not os.path.exists(gui_script_path):
            logging.critical(f"GUI script not found: {gui_script_path}")
            print(f"ERROR: GUI script 'debt_manager_gui.py' not found at {gui_script_path}. Cannot launch application.")
            sys.exit(1)

        # Start the GUI process. We don't wait for it to finish.
        subprocess.Popen([sys.executable, gui_script_path])
        logging.info("Step 3: Debt Management System GUI launched successfully.")
        print("Debt Management System GUI is launching...")
        print("You can close this console window once the GUI appears.")

    except Exception as e:
        logging.critical(f"Step 3: Failed to launch GUI application: {e}", exc_info=True)
        print(f"ERROR: Failed to launch GUI application. Check DebugLog.txt for details. Error: {e}")
        sys.exit(1)

    logging.info("--- Debt Management System Orchestrator finished ---")

if __name__ == "__main__":
    run_orchestrator()
