# SyncAndLaunch.ps1
# Purpose: Orchestrates the Debt Management System:
#          1. Initializes the Access database (creates if missing, ensures tables exist).
#          2. Creates/updates the Excel dashboard template.
#          3. Launches the main Debt Management System UI.
# Deploy in: C:\DebtTracker
# Version: 1.6 (2025-07-16) - Added verbose logging and explicit success checks for each step.
#                            - Added a final logging block to ensure completion status is always recorded.

param ()

$logDir = 'C:\DebtTracker\Logs'
$logFile = "$logDir\SyncLog.txt"

function Write-SyncLog {
    param ([Parameter(Mandatory)][string]$Message)
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message" | Out-File -FilePath $logFile -Append -Force
}

Write-SyncLog 'Starting SyncAndLaunch.ps1'

try {
    # Load System.Windows.Forms assembly for MessageBox.Show calls within this script
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-SyncLog 'System.Windows.Forms assembly loaded for SyncAndLaunch.ps1.'

    $dbInitScriptPath = 'C:\DebtTracker\InitializeDebtDatabase.ps1'
    $excelTemplateScriptPath = 'C:\DebtTracker\CreateExcelTemplate.ps1'
    $uiScriptPath = 'C:\DebtTracker\DebtManagerUI.ps1'
    $excelPath = 'C:\DebtTracker\DebtDashboard.xlsx' # Defined here for checks

    # --- Step 1: Ensure Database is Initialized ---
    Write-SyncLog "Step 1: Running database initialization script: $dbInitScriptPath"
    if (-not (Test-Path $dbInitScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("Database initialization script not found: $dbInitScriptPath", 'Error', 'OK', 'Error')
        exit 1
    }
    . $dbInitScriptPath # Execute the database initialization script
    if ($LASTEXITCODE -ne 0) {
        [System.Windows.Forms.MessageBox]::Show("Database initialization failed. Check DebugLog.txt for details.", 'Error', 'OK', 'Error')
        exit 1
    }
    Write-SyncLog 'Step 1: Database initialization completed successfully.'

    # --- Step 2: Robustly close any running Excel processes and COM objects ---
    Write-SyncLog 'Step 2: Attempting robust closure of Excel processes and COM objects...'
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object {
        Write-SyncLog "Stopping Excel process with ID $($_.Id)."
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
    # Release any lingering COM objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2 # Give more time for processes/objects to release
    Write-SyncLog 'Step 2: Excel process and COM object termination attempt completed.'

    # --- Step 3: Create/Update Excel Template ---
    Write-SyncLog "Step 3: Running Excel template creation script: $excelTemplateScriptPath"
    if (-not (Test-Path $excelTemplateScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("Excel template creation script not found: $excelTemplateScriptPath", 'Error', 'OK', 'Error')
        exit 1
    }
    try {
        . $excelTemplateScriptPath # Execute the Excel template creation script
        if ($LASTEXITCODE -ne 0) {
            [System.Windows.Forms.MessageBox]::Show("Excel template creation failed. Check DebugLog.txt for details.", 'Error', 'OK', 'Error')
            exit 1
        }
        Write-SyncLog 'Step 3: Excel template creation completed successfully.'
    }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-SyncLog "Step 3: Error during Excel template creation script execution: ${actualErrorMessage}"
        [System.Windows.Forms.MessageBox]::Show("Error during Excel template creation: ${actualErrorMessage}`nCheck DebugLog.txt for details.", 'Error', 'OK', 'Error')
        exit 1
    }


    # --- Step 4: Robustly close any running Excel processes and COM objects again before UI launch ---
    Write-SyncLog 'Step 4: Attempting robust closure of Excel processes and COM objects before UI launch...'
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object {
        Write-SyncLog "Stopping Excel process with ID $($_.Id)."
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2 # Give more time for processes/objects to release
    Write-SyncLog 'Step 4: Second Excel process and COM object termination attempt completed.'

    # --- Step 5: Launch Main UI Script ---
    Write-SyncLog "Step 5: Launching DebtManagerUI.ps1: $uiScriptPath"
    if (-not (Test-Path $uiScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("UI script not found: $uiScriptPath", 'Error', 'OK', 'Error')
        exit 1
    }
    # Start-Process is used to run the UI script in a new PowerShell window,
    # so it doesn't block the current script, and provides a separate process for the GUI.
    Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -File `"$uiScriptPath`""
    Write-SyncLog 'Step 5: DebtManagerUI.ps1 launched successfully.'

}
catch {
    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
    Write-SyncLog "CRITICAL ERROR: Startup error in SyncAndLaunch.ps1: ${actualErrorMessage}"
    [System.Windows.Forms.MessageBox]::Show("CRITICAL ERROR: Startup error in SyncAndLaunch.ps1: ${actualErrorMessage}", 'Error', 'OK', 'Error')
}
finally {
    Write-SyncLog 'SyncAndLaunch.ps1 execution completed (or terminated due to error).'
}
