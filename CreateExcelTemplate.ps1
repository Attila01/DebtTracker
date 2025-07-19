# CreateExcelTemplate.ps1
# Purpose: Generate Excel template for Debt Management System
# Deploy in: C:\DebtTracker
# Version: 2.10 (2025-07-16) - PowerShell 7 compatible, fixed InvalidVariableReferenceWithDrive, robust COM handling

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$TestMode
)

# Configuration
$ExcelPath = 'C:\DebtTracker\DebtDashboard.xlsx'
$LogPath = Join-Path -Path 'C:\DebtTracker' -ChildPath 'Logs\DebugLog.txt'
$BasePath = 'C:\DebtTracker'
$LogDir = Join-Path -Path $BasePath -ChildPath 'Logs'
if (-not (Test-Path -Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}
$DebugPreference = if ($TestMode) { 'Continue' } else { 'SilentlyContinue' }

function Write-DebugLog {
    param (
        [Parameter(Mandatory)]
        [string]$Message
    )
    Write-Debug -Message $Message
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}

# Verify Windows environment
if (-not $IsWindows) {
    Write-DebugLog -Message 'COM objects not supported on non-Windows platforms'
    throw 'This script requires Windows for Excel COM automation'
}

# Load assemblies
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-DebugLog -Message 'Assemblies loaded'
}
catch {
    Write-DebugLog -Message "Assembly load error: $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show("Assembly load error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    exit 1
}

function Test-ExcelOpen {
    try {
        return [bool](Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue)
    }
    catch {
        Write-DebugLog -Message "Excel process check error: $($_.Exception.Message)"
        return $false
    }
}

function New-ExcelTemplate {
    [CmdletBinding(SupportsShouldProcess)]
    param ()
    try {
        if (-not (Test-Path -Path $BasePath)) {
            $errorMsg = "Invalid path: $BasePath"
            Write-DebugLog -Message $errorMsg
            throw $errorMsg
        }
        if (Test-ExcelOpen) {
            $errorMsg = 'Excel is open, cannot create template'
            Write-DebugLog -Message $errorMsg
            throw $errorMsg
        }

        $schemas = @{
            'Debts'     = @('DebtID', 'Creditor', 'Amount', 'MinimumPayment', 'SnowballPayment', 'InterestRate', 'DueDate', 'Status')
            'Accounts'  = @('AccountID', 'AccountName', 'Balance', 'AccountType', 'Status')
            'Payments'  = @('PaymentID', 'DebtID', 'Amount', 'PaymentDate', 'PaymentMethod', 'Category')
            'Goals'     = @('GoalID', 'GoalName', 'TargetAmount', 'CurrentAmount', 'TargetDate', 'Status', 'Notes')
            'Assets'    = @('AssetID', 'AssetName', 'Value', 'Category', 'Status')
            'Revenue'   = @('RevenueID', 'Amount', 'DateReceived', 'Source', 'AllocatedTo', 'AllocationType')
            'Categories'= @('CategoryID', 'CategoryName')
        }

        if ($PSCmdlet.ShouldProcess($ExcelPath, 'Create Excel template')) {
            $excel = $null
            $workbook = $null
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $workbook = $excel.Workbooks.Add()

                # Ensure at least one sheet remains
                $defaultSheets = $workbook.Worksheets.Count
                foreach ($sheet in $workbook.Worksheets | Where-Object { $_.Name -notin $schemas.Keys }) {
                    if ($workbook.Worksheets.Count -gt 1) {
                        $sheet.Delete()
                    }
                }

                # Create sheets
                foreach ($tableName in $schemas.Keys) {
                    $worksheet = $workbook.Worksheets | Where-Object { $_.Name -eq $tableName }
                    if (-not $worksheet) {
                        $worksheet = $workbook.Worksheets.Add([System.Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
                        $worksheet.Name = $tableName
                    }
                    $columns = $schemas[$tableName]
                    for ($i = 1; $i -le $columns.Count; $i++) {
                        $worksheet.Cells.Item(1, $i) = $columns[$i - 1]
                    }
                    $worksheet.Rows(1).Font.Bold = $true
                    $worksheet.Columns.AutoFit()
                    Write-DebugLog -Message "Created sheet: $tableName"
                }

                # Save and cleanup
                if (Test-Path -Path $ExcelPath) {
                    Remove-Item -Path $ExcelPath -Force -ErrorAction SilentlyContinue
                    Write-DebugLog -Message "Deleted existing: $ExcelPath"
                }
                $workbook.SaveAs($ExcelPath)
                Write-DebugLog -Message "Template created: $ExcelPath"
                [System.Windows.Forms.MessageBox]::Show("Template created: $ExcelPath", 'Success', 'OK', 'Information') | Out-Null
            }
            finally {
                if ($workbook) {
                    $workbook.Close($false)
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
                }
                if ($excel) {
                    $excel.Quit()
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
    }
    catch {
        Write-DebugLog -Message "Template creation error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Template creation error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        throw
    }
}

# Main execution
try {
    Write-DebugLog -Message 'Starting CreateExcelTemplate.ps1'
    New-ExcelTemplate
}
catch {
    Write-DebugLog -Message "Startup error: $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show("Startup error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    exit 1
}
