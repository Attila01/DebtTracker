# Debt Management System PowerShell Script
# Requires: Microsoft Access Database Engine
# Deploy in: C:\DebtTracker
# Purpose: Manage debts, accounts, payments, goals with GUI, reminders, projections, and Excel syncing
# Version: 3.51 (2025-07-16) - Corrected DDL statements for Access compatibility (VARCHAR to TEXT).
#                               - Added explicit DDL statements for all tables.

param () # No parameters needed for DbPath configuration

# Configuration
$script:config = @{
    DbPath       = 'C:\DebtTracker\db\DebtManager.accdb' # Always use DebtManager.accdb
    ExcelPath    = 'C:\DebtTracker\DebtDashboard.xlsx'
    ReportPath   = 'C:\DebtTracker\reports'
    ConnString   = $null
    CsvPath      = $null
    # Centralized table schemas - defines columns for UI and Sync
    TableSchemas = @{
        'Debts'      = @('DebtID', 'Creditor', 'Amount', 'MinimumPayment', 'SnowballPayment', 'InterestRate', 'DueDate', 'Status')
        'Accounts'   = @('AccountID', 'AccountName', 'Balance', 'AccountType', 'Status')
        'Payments'   = @('PaymentID', 'DebtID', 'Amount', 'PaymentDate', 'PaymentMethod', 'Category')
        'Goals'      = @('GoalID', 'GoalName', 'TargetAmount', 'CurrentAmount', 'TargetDate', 'Status', 'Notes')
        'Assets'     = @('AssetID', 'AssetName', 'Value', 'Category', 'Status')
        'Revenue'    = @('RevenueID', 'Amount', 'DateReceived', 'Source', 'AllocatedTo', 'AllocationType')
        'Categories' = @('CategoryID', 'CategoryName')
    }
}
$dbPath = ${script:config}.DbPath
$script:config.ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${dbPath};"
$script:config.CsvPath = "${script:config}.ReportPath\DebtReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
if (-not (Test-Path 'C:\DebtTracker\db')) { New-Item -ItemType Directory -Path 'C:\DebtTracker\db' -Force | Out-Null }
if (-not (Test-Path ${script:config}.ReportPath)) { New-Item -ItemType Directory -Path ${script:config}.ReportPath -Force | Out-Null }

# Logging
$DebugPreference = 'SilentlyContinue' # Default to SilentlyContinue; can be overridden for debugging
function Write-DebugLog {
    param ([string]$Message)
    Write-Debug ${Message}
    $logFile = 'C:\DebtTracker\Logs\DebugLog.txt'
    if (-not (Test-Path 'C:\DebtTracker\Logs')) { New-Item -ItemType Directory -Path 'C:\DebtTracker\Logs' -Force | Out-Null }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): ${Message}" | Out-File -FilePath ${logFile} -Append -Force
}

# Load assemblies
try {
    Add-Type -AssemblyName System.Windows.Forms, System.Data, System.Drawing -ErrorAction Stop
    Write-DebugLog 'Assemblies loaded'
}
catch {
    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
    Write-DebugLog "Assembly load error: ${actualErrorMessage}"
    [System.Windows.Forms.MessageBox]::Show("Assembly load error: ${actualErrorMessage}`nInstall Microsoft Access Database Engine.", 'Error', 'OK', 'Error') | Out-Null
    exit
}

# Excel utility function (from CreateExcelTemplate.ps1)
function Test-ExcelOpen {
    try { return [bool](Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-DebugLog "Excel process check error: ${actualErrorMessage}"
        return $false
    }
}

# Database and Excel sync
function Sync-Data {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [ValidateSet('ExcelToAccess', 'AccessToExcel')][string]$Direction
    )
    if ($PSCmdlet.ShouldProcess("data", "synchronize ${Direction}")) {
        try {
            if (Test-ExcelOpen) {
                Write-DebugLog 'Excel is open, skipping sync'
                [System.Windows.Forms.MessageBox]::Show('Close Excel to sync data.', 'Warning', 'OK', 'Warning') | Out-Null
                return
            }
            if (-not (Test-Path ${script:config}.ExcelPath)) {
                Write-DebugLog "Excel file missing: ${script:config}.ExcelPath"
                [System.Windows.Forms.MessageBox]::Show("Excel file missing: ${script:config}.ExcelPath`nRun CreateExcelTemplate.ps1.", 'Error', 'OK', 'Error') | Out-Null
                return
            }

            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $excelConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${script:config}.ExcelPath;Extended Properties='Excel 12.0 Xml;HDR=YES;'"
            $excelConn = New-Object System.Data.OleDb.OleDbConnection $excelConnString
            $excelConn.Open()

            foreach ($tableName in ${script:config}.TableSchemas.Keys) {
                $columns = ${script:config}.TableSchemas[$tableName]
                $fieldNames = $columns -join ','
                $placeholders = ('?' * $columns.Count) -join ','
                $table = New-Object System.Data.DataTable
                if ($Direction -eq 'ExcelToAccess') {
                    $excelCmd = $excelConn.CreateCommand()
                    $excelCmd.CommandText = "SELECT * FROM [${tableName}$]"
                    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $excelCmd
                    $adapter.Fill($table) | Out-Null
                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = "DELETE FROM [${tableName}]"
                    $cmd.ExecuteNonQuery() | Out-Null
                    foreach ($row in $table.Rows) {
                        $cmd.CommandText = "INSERT INTO [${tableName}] (${fieldNames}) VALUES (${placeholders})"
                        $cmd.Parameters.Clear()
                        foreach ($column in $columns) {
                            $paramValue = $row[$column] ?? [DBNull]::Value
                            $cmd.Parameters.AddWithValue("@p", $paramValue) | Out-Null
                        }
                        $cmd.ExecuteNonQuery() | Out-Null
                    }
                }
                else { # AccessToExcel
                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = "SELECT * FROM [${tableName}]"
                    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
                    $adapter.Fill($table) | Out-Null
                    $excelCmd = $excelConn.CreateCommand()
                    $excelCmd.CommandText = "DELETE * FROM [${tableName}$]"
                    $excelCmd.ExecuteNonQuery() | Out-Null
                    foreach ($row in $table.Rows) {
                        $excelCmd.CommandText = "INSERT INTO [${tableName}$] (${fieldNames}) VALUES (${placeholders})"
                        $excelCmd.Parameters.Clear()
                        foreach ($column in $columns) {
                            $paramValue = $row[$column] ?? [DBNull]::Value
                            $excelCmd.Parameters.AddWithValue("@p", $paramValue) | Out-Null
                        }
                        $excelCmd.ExecuteNonQuery() | Out-Null
                    }
                }
            }
            Write-DebugLog "${Direction} sync completed"
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Sync error (${Direction}): ${actualErrorMessage}"
            [System.Windows.Forms.MessageBox]::Show("Sync error (${Direction}): ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
        }
        finally {
            if ($excelConn?.State -eq 'Open') { $excelConn.Close() }
            if ($conn?.State -eq 'Open') { $conn.Close() }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

# Database initialization
function Initialize-Database {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ()
    if ($PSCmdlet.ShouldProcess("database", "initialize")) {
        try {
            Write-DebugLog "Initializing database: ${dbPath}"
            if (-not (Test-Path ${dbPath})) {
                Write-DebugLog "Database file not found. Attempting to create: ${dbPath}"
                $adox = New-Object -ComObject ADOX.Catalog
                $adox.Create(${script:config}.ConnString) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($adox) | Out-Null
                Write-DebugLog "Database file created successfully: ${dbPath}"
            }
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $cmd = $conn.CreateCommand()

            # Define DDL statements based on the centralized schemas
            # Changed VARCHAR(X) to TEXT(X) or just TEXT for Access compatibility
            $ddlStatements = @{
                'Debts'      = 'CREATE TABLE Debts (DebtID AUTOINCREMENT PRIMARY KEY, Creditor TEXT(255) NOT NULL, Amount CURRENCY NOT NULL, MinimumPayment CURRENCY, SnowballPayment CURRENCY, InterestRate DOUBLE, DueDate DATETIME, Status TEXT(50))'
                'Accounts'   = 'CREATE TABLE Accounts (AccountID AUTOINCREMENT PRIMARY KEY, AccountName TEXT(255) NOT NULL, Balance CURRENCY NOT NULL, AccountType TEXT(50), Status TEXT(50))'
                'Payments'   = 'CREATE TABLE Payments (PaymentID AUTOINCREMENT PRIMARY KEY, DebtID INTEGER, Amount CURRENCY NOT NULL, PaymentDate DATETIME, PaymentMethod TEXT(255), Category TEXT(50))'
                'Goals'      = 'CREATE TABLE Goals (GoalID AUTOINCREMENT PRIMARY KEY, GoalName TEXT(255) NOT NULL, TargetAmount CURRENCY NOT NULL, CurrentAmount CURRENCY, TargetDate DATETIME, Status TEXT(50), Notes TEXT(255))'
                'Assets'     = 'CREATE TABLE Assets (AssetID AUTOINCREMENT PRIMARY KEY, AssetName TEXT(255) NOT NULL, Value CURRENCY NOT NULL, Category TEXT(50), Status TEXT(50))'
                'Revenue'    = 'CREATE TABLE Revenue (RevenueID AUTOINCREMENT PRIMARY KEY, Amount CURRENCY NOT NULL, DateReceived DATETIME, Source TEXT(255), AllocatedTo INTEGER, AllocationType TEXT(50))'
                'Categories' = 'CREATE TABLE Categories (CategoryID AUTOINCREMENT PRIMARY KEY, CategoryName TEXT(255) NOT NULL)'
            }

            foreach ($tableName in ${script:config}.TableSchemas.Keys) {
                try {
                    if (-not (Test-TableExist -TableName ${tableName})) {
                        $ddl = $ddlStatements[$tableName]
                        if ($null -ne $ddl) {
                            $cmd.CommandText = $ddl
                            $cmd.ExecuteNonQuery() | Out-Null
                            Write-DebugLog "Table ${tableName} created"
                        } else {
                            Write-DebugLog "No DDL statement found for table ${tableName}."
                        }
                    }
                }
                catch {
                    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
                    Write-DebugLog "Error creating table ${tableName}: ${actualErrorMessage}"
                    [System.Windows.Forms.MessageBox]::Show("Error creating table ${tableName}: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
                }
            }
            $conn.Close()
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Database initialization failed: ${actualErrorMessage}"
            # Enhanced message to guide user on missing Access Database Engine
            [System.Windows.Forms.MessageBox]::Show("Database initialization failed: ${actualErrorMessage}`nThis often means the Microsoft Access Database Engine is not installed or is corrupted. Please ensure 'Microsoft Access Database Engine 2010 Redistributable' (or newer) is installed and matches your Office bitness (32-bit vs 64-bit).", 'Error', 'OK', 'Error') | Out-Null
        }
    }
}

function Test-TableExist {
    param ([string]$TableName)
    try {
        $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
        $conn.Open()
        $exists = [bool]($conn.GetSchema('Tables').Rows | Where-Object { $null -ne $_['TABLE_NAME'] -and $_['TABLE_NAME'] -eq ${TableName} })
        $conn.Close()
        return $exists
    }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-DebugLog "Table check error (${TableName}): ${actualErrorMessage}"
        return $false
    }
}

# Data operations
function Get-TableData {
    param (
        [System.Windows.Forms.DataGridView]$GridView,
        [string]$TableName
    )
    try {
        if (-not (Test-TableExist -TableName ${TableName})) { throw "Table ${TableName} missing" }
        $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT * FROM [${TableName}]"
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
        $table = New-Object System.Data.DataTable
        $rows = $adapter.Fill($table)
        $GridView.DataSource = $table
        $GridView.AutoResizeColumns()
        $conn.Close()
        Write-DebugLog "Loaded ${TableName} (${rows} rows)"
    }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-DebugLog "Load error (${TableName}): ${actualErrorMessage}"
        [System.Windows.Forms.MessageBox]::Show("Load error (${TableName}): ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
    }
}

function Update-TableData {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [System.Windows.Forms.DataGridView]$GridView,
        [string]$TableName
    )
    if ($PSCmdlet.ShouldProcess("data in table '$TableName'", "update")) {
        try {
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $adapter = New-Object System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [${TableName}]", $conn)
            $cmdBuilder = New-Object System.Data.OleDb.OleDbCommandBuilder $adapter
            $adapter.UpdateCommand = $cmdBuilder.GetUpdateCommand()
            $adapter.InsertCommand = $cmdBuilder.GetInsertCommand()
            $adapter.DeleteCommand = $cmdBuilder.GetDeleteCommand()
            $adapter.Update($GridView.DataSource) | Out-Null
            $conn.Close()
            [System.Windows.Forms.MessageBox]::Show("Saved ${TableName}.", 'Success', 'OK', 'Information') | Out-Null
            Write-DebugLog "Saved ${TableName}"
            if ($TableName -in @('Accounts', 'Payments', 'Revenue')) { Update-AccountBalances }
            Sync-Data -Direction 'AccessToExcel'
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Save error (${TableName}): ${actualErrorMessage}"
            [System.Windows.Forms.MessageBox]::Show("Save error (${TableName}): ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
        }
    }
}

function Remove-SelectedRow {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [System.Windows.Forms.DataGridView]$GridView,
        [string]$TableName,
        [string]$PrimaryKey
    )
    if ($PSCmdlet.ShouldProcess("selected rows from table '$TableName'", "delete")) {
        try {
            if ($GridView.SelectedRows.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show('No rows selected.', 'Warning', 'OK', 'Warning') | Out-Null
                return
            }
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $cmd = $conn.CreateCommand()
            foreach ($row in $GridView.SelectedRows) {
                $cmd.CommandText = "DELETE FROM [${TableName}] WHERE ${PrimaryKey} = ?"
                $cmd.Parameters.Clear()
                $cmd.Parameters.AddWithValue('@p1', $row.DataBoundItem[${PrimaryKey}]) | Out-Null
                $cmd.ExecuteNonQuery() | Out-Null
            }
            $conn.Close()
            Get-TableData -GridView $GridView -TableName ${TableName}
            if ($TableName -in @('Accounts', 'Payments', 'Revenue')) { Update-AccountBalances }
            Sync-Data -Direction 'AccessToExcel'
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Delete error (${TableName}): ${actualErrorMessage}"
            [System.Windows.Forms.MessageBox]::Show("Delete error (${TableName}): ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
        }
    }
}

function New-Record {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        [string]$TableName,
        [array]$Fields,
        [string]$FormTitle,
        [System.Windows.Forms.DataGridView]$GridView
    )
    try {
        if ($PSCmdlet.ShouldProcess("new record in table '$TableName'", "create")) {
            foreach ($field in $Fields) {
                if (($null -eq ${field}.Name) -or ($null -eq ${field}.Type)) { throw "Invalid field: $($field | ConvertTo-Json)" }
            }
            $form = New-Object System.Windows.Forms.Form -Property @{Text=${FormTitle}; Size=[System.Drawing.Size]::new(400, 400); StartPosition='CenterScreen'}
            $controls = @{}
            $y = 20
            foreach ($field in $Fields) {
                $label = New-Object System.Windows.Forms.Label -Property @{Text=${field}.Name; Location=[System.Drawing.Point]::new(20, $y); Size=[System.Drawing.Size]::new(100, 20)}
                $form.Controls.Add($label)
                if (${field}.Name -match 'Status|AccountType|Category|AllocationType') {
                    $comboBox = New-Object System.Windows.Forms.ComboBox -Property @{Location=[System.Drawing.Point]::new(130, $y); Size=[System.Drawing.Size]::new(200, 20)}
                    if (${field}.Name -eq 'Category') {
                        $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
                        $conn.Open()
                        $cmd = $conn.CreateCommand()
                        $cmd.CommandText = 'SELECT CategoryName FROM Categories'
                        $reader = $cmd.ExecuteReader()
                        while ($reader.Read()) { $comboBox.Items.Add($reader['CategoryName']) }
                        $reader.Close()
                        $conn.Close()
                    }
                    else { $comboBox.Items.AddRange(${field}.Options) }
                    $comboBox.SelectedIndex = 0
                    $form.Controls.Add($comboBox)
                    $controls[${field}.Name] = $comboBox
                }
                elseif (${field}.Name -in @('DebtID', 'AllocatedTo')) {
                    $comboBox = New-Object System.Windows.Forms.ComboBox -Property @{Location=[System.Drawing.Point]::new(130, $y); Size=[System.Drawing.Size]::new(200, 20)}
                    $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
                    $conn.Open()
                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = 'SELECT DebtID, Creditor FROM Debts UNION SELECT AccountID, AccountName FROM Accounts'
                    $reader = $cmd.ExecuteReader()
                    $comboBox.Items.Add('0 - None')
                    while ($reader.Read()) { $comboBox.Items.Add("$($reader[0]) - $($reader[1])") }
                    $reader.Close()
                    $conn.Close()
                    $comboBox.SelectedIndex = 0
                    $form.Controls.Add($comboBox)
                    $controls[${field}.Name] = $comboBox
                }
                else {
                    $textBox = New-Object System.Windows.Forms.TextBox -Property @{Location=[System.Drawing.Point]::new(130, $y); Size=[System.Drawing.Size]::new(200, 20)}
                    $form.Controls.Add($textBox)
                    $controls[${field}.Name] = $textBox
                }
                $y += 30
            }
            $addBtn = New-Object System.Windows.Forms.Button -Property @{Text='Add'; Location=[System.Drawing.Point]::new(130, $y)}
            $addBtn.Add_Click({
                try {
                    $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
                    $conn.Open()
                    $cmd = $conn.CreateCommand()
                    $fieldNames = ($Fields.Name) -join ','
                    $placeholders = ('?' * $Fields.Count) -join ','
                    $cmd.CommandText = "INSERT INTO [${TableName}] (${fieldNames}) VALUES (${placeholders})"
                    $cmd.Parameters.Clear()
                    foreach ($field in $Fields) {
                        $value = $controls[${field}.Name].Text
                        if (($null -eq $value) -or ([string]::IsNullOrWhiteSpace($value)) -and (${field}.Type -notin @('Text', 'Date'))) { throw "Field ${field}.Name cannot be empty" }
                        if (${field}.Name -in @('DebtID', 'AllocatedTo')) { $value = ($value -split ' - ')[0] }
                        if ([string]::IsNullOrWhiteSpace($value)) { $cmd.Parameters.AddWithValue('@p', [DBNull]::Value) }
                        elseif ((${field}.Type -eq 'Date') -and (-not ([DateTime]::TryParse($value, [ref]$null)))) { throw "Invalid date for ${field}.Name" }
                        elseif ((${field}.Type -eq 'Decimal') -and (-not ([decimal]::TryParse($value, [ref]$null)))) { throw "Invalid decimal for ${field}.Name" }
                        else { $cmd.Parameters.AddWithValue('@p', $value) }
                    }
                    $cmd.ExecuteNonQuery() | Out-Null
                    $conn.Close()
                    $form.Close()
                    Get-TableData -GridView $GridView -TableName ${TableName}
                    if (${TableName} -in @('Accounts', 'Payments', 'Revenue')) { Update-AccountBalances }
                    Sync-Data -Direction 'AccessToExcel'
                }
                catch {
                    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
                    Write-DebugLog "Add record error (${TableName}): ${actualErrorMessage}"
                    [System.Windows.Forms.MessageBox]::Show("Add record error (${TableName}): ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
                }
            })
            $form.Controls.Add($addBtn)
            $form.ShowDialog() | Out-Null
        }
    }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-DebugLog "New-Record error (${TableName}): ${actualErrorMessage}"
        [System.Windows.Forms.MessageBox]::Show("New-Record error (${TableName}): ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
    }
}

function Update-AccountBalances {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Descriptive name')]
    param ()
    if ($PSCmdlet.ShouldProcess("all account balances", "update")) {
        try {
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = 'SELECT AccountID FROM Accounts'
            $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
            $accounts = New-Object System.Data.DataTable
            $adapter.Fill($accounts) | Out-Null
            foreach ($account in $accounts.Rows) {
                $accountId = ${account}.AccountID
                $cmd.CommandText = 'SELECT SUM(Amount) FROM Revenue WHERE AllocatedTo = ? AND AllocationType = ?'
                $cmd.Parameters.Clear()
                $cmd.Parameters.AddWithValue('@p1', ${accountId})
                $cmd.Parameters.AddWithValue('@p2', 'Account')
                $deposits = $cmd.ExecuteScalar() ?? 0

                $cmd.CommandText = 'SELECT SUM(Amount) FROM Payments WHERE DebtID = ?'
                $cmd.Parameters.Clear()
                $cmd.Parameters.AddWithValue('@p1', ${accountId})
                $withdrawals = $cmd.ExecuteScalar() ?? 0

                $cmd.CommandText = 'UPDATE Accounts SET Balance = ? WHERE AccountID = ?'
                $cmd.Parameters.Clear()
                $cmd.Parameters.AddWithValue('@p1', $deposits - $withdrawals)
                $cmd.Parameters.AddWithValue('@p2', ${accountId})
                $cmd.ExecuteNonQuery() | Out-Null
            }
            $conn.Close()
            Write-DebugLog 'Account balances updated'
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Balance update error: ${actualErrorMessage}"
            [System.Windows.Forms.MessageBox]::Show("Balance update error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
        }
    }
}

function New-FinancialProjection {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ()
    if ($PSCmdlet.ShouldProcess("financial projection", "generate")) {
        try {
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT SUM(Amount) FROM Debts WHERE Status NOT IN ('Paid Off', 'Closed')"
            $totalDebt = $cmd.ExecuteScalar() ?? 0

            $cmd.CommandText = "SELECT SUM(Balance) FROM Accounts WHERE Status IN ('Open', 'Current', 'Active')"
            $totalSavings = $cmd.ExecuteScalar() ?? 0

            $cmd.CommandText = 'SELECT SUM(Amount) FROM Revenue WHERE DateReceived >= ?'
            $cmd.Parameters.AddWithValue('@p1', (Get-Date).AddMonths(-12))
            $annualIncome = $cmd.ExecuteScalar() ?? 0

            $cmd.CommandText = 'SELECT Amount, MinimumPayment, SnowballPayment FROM Debts WHERE Status NOT IN (''Paid Off'', ''Closed'') ORDER BY Amount ASC'
            $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
            $debts = New-Object System.Data.DataTable
            $adapter.Fill($debts) | Out-Null
            $years = @(3, 5, 7, 10)
            $projections = foreach ($year in $years) {
                $months = $year * 12
                $remainingDebt = $totalDebt
                $snowball = 0
                foreach ($debt in $debts.Rows) {
                    $minimumPayment = $debt.MinimumPayment ?? 0
                    # Removed unused variable 'currentSnowball'

                    $monthlyPayment = $minimumPayment + $snowball
                    $debtAmount = $debt.Amount ?? 0

                    $debtPaid = [Math]::Min($debtAmount, $monthlyPayment * $months)
                    $remainingDebt -= $debtPaid
                    if ($debtPaid -ge $debtAmount) { $snowball += $minimumPayment }
                }
                $savings = $totalSavings * [Math]::Pow(1 + 0.05, $year) + ($annualIncome * 0.2 * $year)
                [PSCustomObject]@{Year=$year; DebtRemaining=$remainingDebt; Savings=$savings; NetWorth=$savings - $remainingDebt}
            }
            $conn.Close()
            $projections | Export-Csv -Path ${script:config}.CsvPath -NoTypeInformation -Force
            [System.Windows.Forms.MessageBox]::Show("Projection generated: ${script:config}.CsvPath", 'Success', 'OK', 'Information') | Out-Null
            Write-DebugLog "Projection generated"
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Projection error: ${actualErrorMessage}"
            [System.Windows.Forms.MessageBox]::Show("Projection error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
        }
    }
}

function Add-Tab {
    param (
        [System.Windows.Forms.TabControl]$TabControl,
        [string]$TabName,
        [string]$TableName,
        [array]$Fields,
        [string]$PrimaryKey,
        [scriptblock]$ExtraButtons
    )
    $tab = New-Object System.Windows.Forms.TabPage -Property @{Text=${TabName}}
    $TabControl.Controls.Add($tab)
    $grid = New-Object System.Windows.Forms.DataGridView -Property @{Location=[System.Drawing.Point]::new(10, 40); Size=[System.Drawing.Size]::new(940, 400)}
    $tab.Controls.Add($grid)
    $loadBtn = New-Object System.Windows.Forms.Button -Property @{Text='Load Data'; Location=[System.Drawing.Point]::new(10, 10)}
    $loadBtn.Add_Click({ Get-TableData -GridView $grid -TableName ${TableName} })
    $tab.Controls.Add($loadBtn)
    $addBtn = New-Object System.Windows.Forms.Button -Property @{Text="Add ${TabName}"; Location=[System.Drawing.Point]::new(10, 450)}
    $addBtn.Add_Click({ New-Record -TableName ${TableName} -Fields ${Fields} -FormTitle "Add ${TabName}" -GridView $grid })
    $tab.Controls.Add($addBtn)
    $deleteBtn = New-Object System.Windows.Forms.Button -Property @{Text='Delete'; Location=[System.Drawing.Point]::new(120, 450)}
    $deleteBtn.Add_Click({ Remove-SelectedRow -GridView $grid -TableName ${TableName} -PrimaryKey ${PrimaryKey} })
    $tab.Controls.Add($deleteBtn)
    $saveBtn = New-Object System.Windows.Forms.Button -Property @{Text='Save Changes'; Location=[System.Drawing.Point]::new(230, 450)}
    $saveBtn.Add_Click({ Update-TableData -GridView $grid -TableName ${TableName} })
    $tab.Controls.Add($saveBtn)
    if ($ExtraButtons) {
        & ${ExtraButtons} -Tab $tab -Grid $grid
    }
}

function Show-MainForm {
    try {
        $form = New-Object System.Windows.Forms.Form -Property @{Text='Debt Management System'; Size=[System.Drawing.Size]::new(1000, 600); StartPosition='CenterScreen'}
        $form.FormClosing.Add({ Sync-Data -Direction 'AccessToExcel' })
        $tabControl = New-Object System.Windows.Forms.TabControl -Property @{Location=[System.Drawing.Point]::new(10, 10); Size=[System.Drawing.Size]::new(960, 540)}
        $form.Controls.Add($tabControl)

        # Dashboard
        $dashboardTab = New-Object System.Windows.Forms.TabPage -Property @{Text='Dashboard'}
        $tabControl.Controls.Add($dashboardTab)
        $dashboardGrid = New-Object System.Windows.Forms.DataGridView -Property @{Location=[System.Drawing.Point]::new(10, 40); Size=[System.Drawing.Size]::new(940, 400)}
        $dashboardTab.Controls.Add($dashboardGrid)
        $dashboardLoadBtn = New-Object System.Windows.Forms.Button -Property @{Text='Load Snowball Data'; Location=[System.Drawing.Point]::new(10, 450)}
        $dashboardLoadBtn.Add_Click({
            try {
                $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
                $conn.Open()
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = 'SELECT Creditor, Amount, MinimumPayment, SnowballPayment, Status FROM Debts WHERE Status NOT IN (''Paid Off'', ''Closed'') ORDER BY Amount ASC'
                $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
                $table = New-Object System.Data.DataTable
                $rows = $adapter.Fill($table)
                $dashboardGrid.DataSource = $table
                $dashboardGrid.AutoResizeColumns()
                $conn.Close()
                Write-DebugLog "Snowball data loaded (${rows} rows)"
            }
            catch {
                $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
                Write-DebugLog "Snowball load error: ${actualErrorMessage}"
                [System.Windows.Forms.MessageBox]::Show("Snowball load error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
            }
        })
        $dashboardTab.Controls.Add($dashboardLoadBtn)

        # Other tabs - Ensured schemas match centralized TableSchemas
        Add-Tab -TabControl $tabControl -TabName 'Debts' -TableName 'Debts' -PrimaryKey 'DebtID' -Fields @(
            @{Name='Creditor'; Type='Text'},
            @{Name='Amount'; Type='Decimal'},
            @{Name='MinimumPayment'; Type='Decimal'},
            @{Name='SnowballPayment'; Type='Decimal'},
            @{Name='InterestRate'; Type='Decimal'},
            @{Name='DueDate'; Type='Date'},
            @{Name='Status'; Type='Text'; Options=@('Open', 'Closed', 'Current', 'In Collection', 'Paid Off')}
        )
        Add-Tab -TabControl $tabControl -TabName 'Accounts' -TableName 'Accounts' -PrimaryKey 'AccountID' -Fields @(
            @{Name='AccountName'; Type='Text'},
            @{Name='Balance'; Type='Decimal'},
            @{Name='AccountType'; Type='Text'; Options=@('Checking', 'Savings', 'Credit')},
            @{Name='Status'; Type='Text'; Options=@('Open', 'Closed', 'Current')}
        )
        Add-Tab -TabControl $tabControl -TabName 'Bills' -TableName 'Payments' -PrimaryKey 'PaymentID' -Fields @(
            @{Name='DebtID'; Type='Integer'},
            @{Name='Amount'; Type='Decimal'},
            @{Name='PaymentDate'; Type='Date'},
            @{Name='PaymentMethod'; Type='Text'},
            @{Name='Category'; Type='Text'}
        )
        Add-Tab -TabControl $tabControl -TabName 'Transactions' -TableName 'Payments' -PrimaryKey 'PaymentID' -Fields @(
            @{Name='DebtID'; Type='Integer'},
            @{Name='Amount'; Type='Decimal'},
            @{Name='PaymentDate'; Type='Date'},
            @{Name='PaymentMethod'; Type='Text'},
            @{Name='Category'; Type='Text'}
        )
        Add-Tab -TabControl $tabControl -TabName 'Goals' -TableName 'Goals' -PrimaryKey 'GoalID' -Fields @(
            @{Name='GoalName'; Type='Text'},
            @{Name='TargetAmount'; Type='Decimal'},
            @{Name='CurrentAmount'; Type='Decimal'},
            @{Name='TargetDate'; Type='Date'},
            @{Name='Status'; Type='Text'; Options=@('Planned', 'In Progress', 'Completed')},
            @{Name='Notes'; Type='Text'}
        ) -ExtraButtons {
            param ($Tab, $Grid)
            $updateBtn = New-Object System.Windows.Forms.Button -Property @{Text='Update Progress'; Location=[System.Drawing.Point]::new(340, 450)}
            $updateBtn.Add_Click({ Update-GoalProgress -GridView $Grid })
            $Tab.Controls.Add($updateBtn)
            $projBtn = New-Object System.Windows.Forms.Button -Property @{Text='Generate Projection'; Location=[System.Drawing.Point]::new(450, 450)}
            $projBtn.Add_Click({ New-FinancialProjection })
            $Tab.Controls.Add($projBtn)
        }
        Add-Tab -TabControl $tabControl -TabName 'Assets' -TableName 'Assets' -PrimaryKey 'AssetID' -Fields @(
            @{Name='AssetName'; Type='Text'},
            @{Name='Value'; Type='Decimal'},
            @{Name='Category'; Type='Text'},
            @{Name='Status'; Type='Text'; Options=@('Active', 'Inactive', 'Sold')}
        )
        Add-Tab -TabControl $tabControl -TabName 'Revenue' -TableName 'Revenue' -PrimaryKey 'RevenueID' -Fields @(
            @{Name='Amount'; Type='Decimal'},
            @{Name='DateReceived'; Type='Date'},
            @{Name='Source'; Type='Text'},
            @{Name='AllocatedTo'; Type='Integer'},
            @{Name='AllocationType'; Type='Text'; Options=@('Account', 'Debt', 'Other')}
        )

        # Reports Tab
        $reportsTab = New-Object System.Windows.Forms.TabPage -Property @{Text='Reports'}
        $tabControl.Controls.Add($reportsTab)
        $reportTypeCombo = New-Object System.Windows.Forms.ComboBox -Property @{Location=[System.Drawing.Point]::new(100, 20); Size=[System.Drawing.Size]::new(200, 20)}
        $reportTypeCombo.Items.AddRange(@('Debt Summary', 'Daily Expenses', 'Snowball Progress', 'Account Balances'))
        $reportTypeCombo.SelectedIndex = 0
        $reportsTab.Controls.Add($reportTypeCombo)
        $reportDatePicker = New-Object System.Windows.Forms.DateTimePicker -Property @{Location=[System.Drawing.Point]::new(100, 50); Size=[System.Drawing.Size]::new(200, 20)}
        $reportsTab.Controls.Add($reportDatePicker)
        $reportBtn = New-Object System.Windows.Forms.Button -Property @{Text='Generate Report'; Location=[System.Drawing.Point]::new(100, 80)}
        $reportBtn.Add_Click({
            try {
                $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
                $conn.Open()
                $cmd = $conn.CreateCommand()
                $reportData = New-Object System.Data.DataTable
                $reportType = $reportTypeCombo.SelectedItem
                if ($null -eq ${reportType}) { throw 'No report type selected' }
                switch (${reportType}) {
                    'Debt Summary' { $cmd.CommandText = 'SELECT * FROM Debts' }
                    'Daily Expenses' {
                        if ($null -eq ${reportDatePicker}.Value) { throw 'Invalid date' }
                        $cmd.CommandText = 'SELECT * FROM Payments WHERE PaymentDate = ?'
                        $cmd.Parameters.AddWithValue('@p1', ${reportDatePicker}.Value.Date)
                    }
                    'Account Balances' { $cmd.CommandText = 'SELECT * FROM Accounts' }
                    'Snowball Progress' { $cmd.CommandText = 'SELECT Creditor, Amount, MinimumPayment, SnowballPayment, Status FROM Debts WHERE Status NOT IN (''Paid Off'', ''Closed'') ORDER BY Amount ASC' }
                }
                $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
                $rows = $adapter.Fill($reportData)
                $conn.Close()
                if ($rows -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("No data for ${reportType}.", 'Information', 'OK', 'Information') | Out-Null
                }
                else {
                    $reportData | Export-Csv -Path ${script:config}.CsvPath -NoTypeInformation -Force
                    [System.Windows.Forms.MessageBox]::Show("Report generated: ${script:config}.CsvPath", 'Success', 'OK', 'Information') | Out-Null
                }
            }
            catch {
                $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
                Write-DebugLog "Report error: ${actualErrorMessage}"
                [System.Windows.Forms.MessageBox]::Show("Report error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
            }
        })
        $reportsTab.Controls.Add($reportBtn)

        $form.ShowDialog() | Out-Null
    }
    catch {
        $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
        Write-DebugLog "Main form error: ${actualErrorMessage}"
        [System.Windows.Forms.MessageBox]::Show("Main form error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
    }
}

function Update-GoalProgress {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ([System.Windows.Forms.DataGridView]$GridView)
    if ($PSCmdlet.ShouldProcess("goal progress", "update")) {
        try {
            $conn = New-Object System.Data.OleDb.OleDbConnection ${script:config}.ConnString
            $conn.Open()
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = 'SELECT GoalID, TargetAmount FROM Goals'
            $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
            $goals = New-Object System.Data.DataTable
            $adapter.Fill($goals) | Out-Null
            foreach ($goal in $goals.Rows) {
                $cmd.CommandText = 'SELECT SUM(Amount) FROM Payments WHERE Category = ? AND PaymentDate <= ?'
                $cmd.Parameters.Clear()
                $cmd.Parameters.AddWithValue('@p1', 'Debt Payment')
                $cmd.Parameters.AddWithValue('@p2', (Get-Date))
                $progress = $cmd.ExecuteScalar() ?? 0

                $cmd.CommandText = 'UPDATE Goals SET CurrentAmount = ?, Status = ? WHERE GoalID = ?'
                $cmd.Parameters.Clear()
                $cmd.Parameters.AddWithValue('@p1', $progress)
                $cmd.Parameters.AddWithValue('@p2', $(if ($progress -ge ${goal}.TargetAmount) { 'Completed' } else { 'In Progress' }))
                $cmd.Parameters.AddWithValue('@p3', ${goal}.GoalID)
                $cmd.ExecuteNonQuery() | Out-Null
            }
            $conn.Close()
            Get-TableData -GridView $GridView -TableName 'Goals'
            Sync-Data -Direction 'AccessToExcel'
        }
        catch {
            $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
            Write-DebugLog "Goal progress error: ${actualErrorMessage}"
            [System.Windows.Forms.MessageBox]::Show("Goal progress error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
        }
    }
}

# Main execution
try {
    Initialize-Database
    Sync-Data -Direction 'ExcelToAccess'
    Show-MainForm
}
catch {
    $actualErrorMessage = $_.Exception?.Message ?? $_.ToString()
    Write-DebugLog "Startup error: ${actualErrorMessage}"
    [System.Windows.Forms.MessageBox]::Show("Startup error: ${actualErrorMessage}", 'Error', 'OK', 'Error') | Out-Null
}
