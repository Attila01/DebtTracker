# debt_manager_orchestrator.py
# Purpose: Orchestrates the Debt Management System startup.
#          1. Initializes the SQLite database.
#          2. Creates/updates the Excel dashboard template.
#          3. Performs initial data synchronization from SQLite to Excel.
#          4. Launches the main Python GUI script.
# Deploy in: C:\DebtTracker
# Version: 1.1 (2025-07-19) - Updated to launch debt_manager_gui.py instead of DebtManagerUI.ps1.

import os
import subprocess
import logging
import time # For potential delays

# Define paths to other scripts and the Excel file
# Assuming all scripts are in C:\DebtTracker
BASE_DIR = 'C:\\DebtTracker'
DB_INIT_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_db_init.py')
EXCEL_TEMPLATE_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_excel_template.py')
EXCEL_SYNC_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_excel_sync.py')
# Updated UI script path to the new Python GUI
UI_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_gui.py')
LOG_DIR = os.path.join(BASE_DIR, 'Logs')
LOG_FILE = os.path.join(LOG_DIR, 'OrchestratorLog.txt')

# Ensure log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s: %(message)s',
                    handlers=[
                        logging.FileHandler(LOG_FILE, mode='a'),
                        logging.StreamHandler()
                    ])

def run_python_script(script_path, script_name):
    """Helper function to run a Python script as a subprocess."""
    logging.info(f"Running Python script: {script_name} from {script_path}")
    if not os.path.exists(script_path):
        logging.error(f"Error: {script_name} not found at {script_path}")
        raise FileNotFoundError(f"{script_name} not found.")

    try:
        # Use sys.executable to ensure the correct Python interpreter is used
        # For a simple script, 'python' usually works if it's in PATH
        result = subprocess.run(['python', script_path], check=True, capture_output=True, text=True)
        logging.info(f"{script_name} stdout:\n{result.stdout}")
        if result.stderr:
            logging.warning(f"{script_name} stderr:\n{result.stderr}")
        logging.info(f"{script_name} completed successfully.")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running {script_name} (Exit Code: {e.returncode}): {e.stderr}", exc_info=True)
        raise
    except Exception as e:
        logging.error(f"Unexpected error running {script_name}: {e}", exc_info=True)
        raise

def run_python_gui_script(script_path, script_name):
    """
    Helper function to run a Python GUI script.
    Temporarily uses subprocess.run to capture stderr for debugging.
    """
    logging.info(f"Launching Python GUI script: {script_name} from {script_path}")
    if not os.path.exists(script_path):
        logging.error(f"Error: {script_name} not found at {script_path}")
        raise FileNotFoundError(f"{script_name} not found.")

    try:
        # TEMPORARY CHANGE FOR DEBUGGING: Use subprocess.run to capture output
        # This will block the orchestrator until the GUI script exits.
        # We are using 'python.exe' to ensure any console output is visible.
        result = subprocess.run(['python.exe', script_path], check=True, capture_output=True, text=True)
        logging.info(f"{script_name} stdout:\n{result.stdout}")
        if result.stderr:
            logging.error(f"{script_name} stderr (GUI script error):\n{result.stderr}")
            print(f"GUI Error: Check OrchestratorLog.txt for details. Stderr:\n{result.stderr}") # Print to console for immediate feedback
        logging.info(f"{script_name} completed (or exited).")
    except subprocess.CalledProcessError as e:
        logging.critical(f"GUI script {script_name} crashed (Exit Code: {e.returncode}): {e.stderr}", exc_info=True)
        print(f"CRITICAL GUI ERROR: {script_name} crashed. Check OrchestratorLog.txt for details. Stderr:\n{e.stderr}")
        raise
    except Exception as e:
        logging.critical(f"Unexpected error launching GUI script {script_name}: {e}", exc_info=True)
        print(f"CRITICAL GUI LAUNCH ERROR: Check OrchestratorLog.txt for details. Error: {e}")
        raise

def main():
    """Orchestrates the Debt Management System startup process."""
    logging.info("--- Starting Debt Management System Orchestrator ---")

    try:
        # Step 1: Initialize Database
        logging.info("Step 1: Initializing database...")
        run_python_script(DB_INIT_SCRIPT, 'debt_manager_db_init.py')
        logging.info("Step 1: Database initialized successfully.")

        # Step 2: Create/Update Excel Template
        logging.info("Step 2: Creating/updating Excel template...")
        run_python_script(EXCEL_TEMPLATE_SCRIPT, 'debt_manager_excel_template.py')
        logging.info("Step 2: Excel template created/updated successfully.")

        # Step 3: Perform Initial SQLite to Excel Sync
        logging.info("Step 3: Performing initial SQLite to Excel sync...")
        # We call the sqlite_to_excel function directly from the script
        # by running the script itself, as it has the main execution block.
        run_python_script(EXCEL_SYNC_SCRIPT, 'debt_manager_excel_sync.py')
        logging.info("Step 3: Initial SQLite to Excel sync completed successfully.")

        # Step 4: Launch Main UI Script (Python GUI)
        logging.info("Step 4: Launching Debt Management System GUI...")
        run_python_gui_script(UI_SCRIPT, 'debt_manager_gui.py')
        logging.info("Step 4: Debt Management System GUI launched successfully.")

    except FileNotFoundError as e:
        logging.critical(f"Orchestration failed: {e}. Please ensure all necessary scripts are in {BASE_DIR}.")
        print(f"CRITICAL ERROR: {e}. Please ensure all necessary scripts are in {BASE_DIR}. Check {LOG_FILE} for details.")
    except Exception as e:
        logging.critical(f"CRITICAL ERROR during orchestration: {e}", exc_info=True)
        print(f"CRITICAL ERROR: An unexpected error occurred during orchestration. Check {LOG_FILE} for details.")
    finally:
        logging.info("--- Debt Management System Orchestrator Finished ---")

if __name__ == "__main__":
    main()
