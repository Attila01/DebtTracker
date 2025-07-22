# debt_manager_orchestrator.py
# Purpose: Orchestrates the Debt Management System startup.
#          1. Initializes the SQLite database.
#          2. Performs initial data synchronization from SQLite to CSV files.
#          3. Launches the main Python GUI script.
# Deploy in: C:\DebtTracker
# Version: 1.6 (2025-07-21) - Critical fix for WinError 87 during GUI launch.
#                            Uses a more compatible subprocess.Popen call for Windows Store Python.

import os
import subprocess
import logging
import time
import sys

# Define paths to other scripts and the CSV directory
BASE_DIR = 'C:\\DebtTracker'
DB_INIT_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_db_init.py')
CSV_SYNC_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_csv_sync.py')
UI_SCRIPT = os.path.join(BASE_DIR, 'debt_manager_gui.py')
LOG_DIR = os.path.join(BASE_DIR, 'Logs')
LOG_FILE = os.path.join(LOG_DIR, 'OrchestratorLog.txt')

# --- Determine the correct Python executable path ---
# sys.executable gives the absolute path to the Python interpreter
PYTHON_EXECUTABLE = sys.executable

os.makedirs(LOG_DIR, exist_ok=True)

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
        result = subprocess.run([PYTHON_EXECUTABLE, script_path], check=True, capture_output=True, text=True)
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
    Helper function to launch a Python GUI script in a non-blocking way.
    This version uses a highly compatible method for Windows Store Python distributions.
    """
    logging.info(f"Launching Python GUI script: {script_name} from {script_path}")
    if not os.path.exists(script_path):
        logging.error(f"Error: {script_name} not found at {script_path}")
        raise FileNotFoundError(f"{script_name} not found.")

    try:
        subprocess.Popen([PYTHON_EXECUTABLE, script_path])
        logging.info(f"{script_name} launched (potentially with a new console window).")
        time.sleep(0.5)

    except Exception as e:
        logging.critical(f"CRITICAL ERROR: Failed to launch GUI script {script_name}: {e}", exc_info=True)
        raise

def main():
    """Orchestrates the Debt Management System startup process."""
    logging.info("--- Starting Debt Management System Orchestrator ---")

    try:
        # Step 1: Initialize Database
        logging.info("Step 1: Initializing database...")
        run_python_script(DB_INIT_SCRIPT, 'debt_manager_db_init.py')
        logging.info("Step 1: Database initialized successfully.")

        # Step 2: Perform Initial SQLite to CSV Sync
        logging.info("Step 2: Performing initial SQLite to CSV sync...")
        run_python_script(CSV_SYNC_SCRIPT, 'debt_manager_csv_sync.py')
        logging.info("Step 2: Initial SQLite to CSV sync completed successfully.")

        # Step 3: Launch Main UI Script (Python GUI)
        logging.info("Step 3: Launching Debt Management System GUI...")
        run_python_gui_script(UI_SCRIPT, 'debt_manager_gui.py')
        logging.info("Step 3: Debt Management System GUI launched successfully.")

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