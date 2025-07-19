# InitializeDebtDatabase.ps1
# Purpose: Creates or updates the Debt Management System Access Database (DebtManager.accdb)
# Requires: Microsoft Access Database Engine (64-bit recommended if using 64-bit PowerShell/Office)
# Deploy in: C:\DebtTracker
# Version: 1.4 (2025-07-16) - Final refinement of DDL for 'Assets' table (AssetDetail to AssetDescription, Status to AssetStatus).

param ()

# Configuration for database path and schema
$script:config = @{
    DbPath       = 'C:\DebtTracker\db\DebtManager.accdb' # Always use DebtManager.accdb
    ConnString   = $null
    # Centralized table schemas - defines columns and their DDL for database creation
    TableSchemas = @{
        'Debts'      = 'CREATE TABLE Debts (DebtID AUTOINCREMENT PRIMARY KEY, Creditor TEXT(255) NOT NULL, Amount CURRENCY NOT NULL, MinimumPayment CURRENCY, SnowballPayment CURRENCY, InterestRate DOUBLE, DueDate DATETIME, Status TEXT(50))'
        'Accounts'   = 'CREATE TABLE Accounts (AccountID AUTOINCREMENT PRIMARY KEY, AccountName TEXT(255) NOT NULL, Balance CURRENCY NOT NULL, AccountType TEXT(50), Status TEXT(50))'
        'Payments'   = 'CREATE TABLE Payments (PaymentID AUTOINCREMENT PRIMARY KEY, DebtID INTEGER, Amount CURRENCY NOT NULL, PaymentDate DATETIME, PaymentMethod TEXT(255), Category TEXT(50))'
        'Goals'      = 'CREATE TABLE Goals (GoalID AUTOINCREMENT PRIMARY KEY, GoalName TEXT(255) NOT NULL, TargetAmount CURRENCY NOT NULL, CurrentAmount CURRENCY, TargetDate DATETIME, Status TEXT(50), Notes TEXT(255))'
        'Assets'     = 'CREATE TABLE Assets (AssetID AUTOINCREMENT PRIMARY KEY, AssetName TEXT(255) NOT NULL, Value CURRENCY NOT NULL, AssetDescription TEXT, AssetStatus TEXT)' # Renamed AssetDetail to AssetDescription, Status to AssetStatus
        'Revenue'    = 'CREATE TABLE Revenue (RevenueID AUTOINCREMENT PRIMARY KEY, Amount CURRENCY NOT NULL, DateReceived DATETIME, Source TEXT(255), AllocatedTo INTEGER, AllocationType TEXT(50))'
        'Categories' = 'CREATE TABLE Categories (CategoryID AUTOINCREMENT PRIMARY KEY, CategoryName TEXT(255) NOT NULL)'
    }
}
$dbPath = ${script:config}.DbPath
$script:config.ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${dbPath};"

# Ensure the database directory exists
if (-not (Test-Path 'C:\DebtTracker\db')) { New-Item -ItemType Directory -Path 'C:\DebtTracker\db' -Force | Out-Null }

# Logging function (simplified for this standalone script)
function Write-DebugLog {
    param ([string]$Message)
    # For this standalone script, we'll write to console and log file
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): ${Message}"
    $logFile = 'C:\DebtTracker\Logs\DebugLog.txt'
    if (-not (Test-Path 'C:\DebtTracker\Logs')) { New-Item -ItemType Directory -Path (Split-Path $logFile) -Force | Out-Null } # Ensure log file directory exists
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): ${Message}" | Out-File -FilePath ${logFile} -Append -Force
}

# Load necessary assemblies
try {
    Add-Type -AssemblyName System.Data -ErrorAction Stop
    Write-DebugLog 'System.Data assembly loaded for database operations.'
}
catch {
    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
    Write-DebugLog "Assembly load error: ${actualErrorMessage}"
    Write-Host "ERROR: Assembly load failed. This script requires the Microsoft Access Database Engine. Please ensure it's installed and its bitness matches your PowerShell process (32-bit vs 64-bit)."
    exit 1 # Exit with an error code
}

# Function to check if a table exists in the database
function Test-TableExist {
    param ([string]$TableName)
    try {
        $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
        $conn.Open()
        # GetSchema returns a DataTable with information about tables, views, etc.
        # We filter for TABLE_NAME to check for existence
        $exists = [bool]($conn.GetSchema('Tables').Rows | Where-Object { $null -ne $_['TABLE_NAME'] -and $_['TABLE_NAME'] -eq ${TableName} })
        $conn.Close()
        return $exists
    }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-DebugLog "Table existence check error for ${TableName}: ${actualErrorMessage}"
        return $false
    }
}

# Main database initialization function
function Initialize-Database {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ()
    if ($PSCmdlet.ShouldProcess("database", "initialize")) {
        try {
            Write-DebugLog "Initializing database: ${dbPath}"

            # Create the database file if it doesn't exist
            if (-not (Test-Path ${dbPath})) {
                Write-DebugLog "Database file not found. Attempting to create: ${dbPath}"
                try {
                    $adox = New-Object -ComObject ADOX.Catalog
                    $adox.Create(${script:config}.ConnString) | Out-Null
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($adox) | Out-Null
                    Write-DebugLog "Database file created successfully: ${dbPath}"
                }
                catch {
                    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
                    Write-DebugLog "CRITICAL ERROR: Failed to create database file: ${actualErrorMessage}"
                    Write-Host "CRITICAL ERROR: Failed to create database file: ${actualErrorMessage}`nEnsure the Microsoft Access Database Engine is correctly installed and its bitness matches your PowerShell process (32-bit vs 64-bit)."
                    exit 1 # Exit if file creation fails
                }
            }

            # Open connection to the database
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $cmd = $conn.CreateCommand()

            # Iterate through all defined tables and create them if missing
            foreach ($tableName in ${script:config}.TableSchemas.Keys) {
                try {
                    if (-not (Test-TableExist -TableName ${tableName})) {
                        $ddl = ${script:config}.TableSchemas[$tableName] # Get DDL from centralized schema
                        if ($null -ne $ddl) {
                            $cmd.CommandText = $ddl
                            $cmd.ExecuteNonQuery() | Out-Null
                            Write-DebugLog "Table ${tableName} created successfully."
                        } else {
                            Write-DebugLog "Warning: No DDL statement found for table ${tableName} in TableSchemas."
                        }
                    } else {
                        Write-DebugLog "Table ${tableName} already exists. Skipping creation."
                    }
                }
                catch {
                    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
                    Write-DebugLog "CRITICAL ERROR: Failed to create table ${tableName}: ${actualErrorMessage}"
                    Write-Host "CRITICAL ERROR: Failed to create table ${tableName}: ${actualErrorMessage}`nThis indicates a problem with the DDL or Access Database Engine."
                    # Do not exit here, attempt to create other tables to provide more comprehensive error logging
                }
            }
            Write-DebugLog "Database initialization process completed."
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Overall Database initialization failed: ${actualErrorMessage}"
            Write-Host "Overall Database initialization failed: ${actualErrorMessage}`nReview the DebugLog.txt for specific table creation errors."
            exit 1 # Exit with an error code
        }
        finally {
            if ($conn?.State -eq 'Open') { $conn.Close() }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

# Main execution for this script
try {
    Initialize-Database
}
catch {
    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
    Write-DebugLog "Startup error in InitializeDebtDatabase.ps1: ${actualErrorMessage}"
    Write-Host "Startup error in InitializeDebtDatabase.ps1: ${actualErrorMessage}"
    exit 1 # Exit with an error code
}
