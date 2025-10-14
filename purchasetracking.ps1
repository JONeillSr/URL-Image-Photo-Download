<#
.SYNOPSIS
    Enhanced purchase tracking system with database storage and advanced price extraction.

.DESCRIPTION
    This script automates price research for auction and closeout purchases by searching
    online retailers, storing price history in a database, and providing pricing recommendations.

.PARAMETER InputCSV
    Path to the input CSV file exported from the purchase tracking spreadsheet.

.PARAMETER DatabasePath
    Path to SQLite database file. Creates if doesn't exist.

.PARAMETER UseAPI
    Switch to enable API-based price fetching where available.

.PARAMETER UpdateExisting
    Switch to update existing items in database with latest prices.

.PARAMETER ShowDashboard
    Switch to display the interactive dashboard after processing.

.PARAMETER TestCSV
    Switch to test CSV file structure without processing.

.EXAMPLE
    .\PurchaseTracking.ps1 -InputCSV "purchases.csv" -DatabasePath ".\data.db" -ShowDashboard

.EXAMPLE
    .\PurchaseTracking.ps1 -DatabasePath ".\data.db" -UpdateExisting

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10/07/2025
    Version: 2.1.0
    Change Date: 10/13/2025
    Change Purpose: Refactored to use shared module

.CHANGELOG
    2.1.0 - Refactored to use PurchaseTracking.psm1 shared module
    2.0.0 - Enhanced with database storage and price extraction
    1.0.0 - Initial release
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$InputCSV,

    [Parameter(Mandatory=$false)]
    [string]$DatabasePath = ".\PurchaseTracking.db",

    [Parameter(Mandatory=$false)]
    [switch]$UseAPI,

    [Parameter(Mandatory=$false)]
    [switch]$UpdateExisting,

    [Parameter(Mandatory=$false)]
    [string]$LogPath = $PSScriptRoot,

    [Parameter(Mandatory=$false)]
    [switch]$ShowDashboard,

    [Parameter(Mandatory=$false)]
    [switch]$TestCSV
)

#region Module Requirements

# Import shared module
$modulePath = Join-Path $PSScriptRoot "modules\PurchaseTracking.psm1"
if (-not (Test-Path $modulePath)) {
    Write-Error "Required module not found: $modulePath"
    Write-Error "Please ensure PurchaseTracking.psm1 is in the same directory as this script."
    exit 1
}

try {
    Import-Module $modulePath -Force -ErrorAction Stop
    Write-Verbose "Successfully loaded PurchaseTracking module"
}
catch {
    Write-Error "Failed to import PurchaseTracking module: $_"
    exit 1
}

# Check and install required modules
$requiredModules = @('PSSQLite')
foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing required module: $module" -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "Module $module installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install $module. Please run: Install-Module $module -Scope CurrentUser"
            exit 1
        }
    }
}

try {
    Import-Module PSSQLite -ErrorAction Stop
}
catch {
    Write-Error "Failed to import PSSQLite module. Please ensure it's installed."
    exit 1
}

# Add .NET assemblies
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Net.Http

#endregion

#region Helper Functions

function Test-CSVStructure {
    param([string]$CSVPath)

    Write-Host "`n=== CSV Structure Test ===" -ForegroundColor Cyan

    if (-not (Test-Path $CSVPath)) {
        Write-Error "CSV file not found: $CSVPath"
        return $false
    }

    $data = Import-Csv -Path $CSVPath

    if ($data.Count -eq 0) {
        Write-Warning "CSV file is empty"
        return $false
    }

    $firstItem = $data[0]
    $columns = $firstItem.PSObject.Properties.Name

    Write-Host "Found $($data.Count) items in CSV" -ForegroundColor Green
    Write-Host "Columns found:" -ForegroundColor Yellow
    $columns | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

    # Check required columns
    $requiredColumns = @('Description', 'Per Item Cost')
    $missingRequired = @()

    foreach ($col in $requiredColumns) {
        if ($col -notin $columns) {
            $missingRequired += $col
        }
    }

    if ($missingRequired.Count -gt 0) {
        Write-Host "`nMissing REQUIRED columns:" -ForegroundColor Red
        $missingRequired | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        return $false
    }

    Write-Host "`nAll required columns present!" -ForegroundColor Green
    return $true
}

#endregion

#region Summary Report

function New-SummaryReport {
    param([string]$DatabasePath)

    Write-Host "`n=== Summary Report ===" -ForegroundColor Cyan

    $query = @"
    SELECT
        COUNT(*) as TotalItems,
        SUM(Quantity) as TotalQuantity,
        SUM(PerItemCost * Quantity) as TotalInvestment,
        SUM(SuggestedPrice * Quantity) as PotentialRevenue,
        AVG(CASE WHEN SuggestedPrice > 0 AND PerItemCost > 0
             THEN (SuggestedPrice - PerItemCost) / PerItemCost * 100
             ELSE 0 END) as AvgProfitMargin
    FROM Items WHERE IsActive = 1
"@

    $stats = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query

    if ($stats) {
        Write-Host "Total Items: $($stats.TotalItems)" -ForegroundColor White
        Write-Host "Total Quantity: $($stats.TotalQuantity)" -ForegroundColor White
        Write-Host "Total Investment: `$$([Math]::Round($stats.TotalInvestment, 2))" -ForegroundColor Yellow
        Write-Host "Potential Revenue: `$$([Math]::Round($stats.PotentialRevenue, 2))" -ForegroundColor Green
        Write-Host "Average Profit Margin: $([Math]::Round($stats.AvgProfitMargin, 1))%" -ForegroundColor Green

        if ($stats.TotalInvestment -gt 0) {
            $totalProfit = $stats.PotentialRevenue - $stats.TotalInvestment
            Write-Host "Potential Profit: `$$([Math]::Round($totalProfit, 2))" -ForegroundColor Cyan
        }
    }
}

#endregion

#region Main Execution

try {
    Write-Host "`n=== Purchase Tracking System v2.1.0 ===" -ForegroundColor Cyan
    Write-Host "Author: John O'Neill Sr. | Company: Azure Innovators`n" -ForegroundColor Gray

    # Test CSV mode
    if ($TestCSV) {
        if (-not $InputCSV) {
            Write-Error "Please specify -InputCSV when using -TestCSV"
            exit 1
        }

        $testResult = Test-CSVStructure -CSVPath $InputCSV
        if ($testResult) {
            Write-Host "`nCSV is ready for processing!" -ForegroundColor Green
        }
        else {
            Write-Host "`nCSV needs adjustments." -ForegroundColor Red
        }
        exit 0
    }

    # Validate parameters
    if (-not $InputCSV -and -not $UpdateExisting) {
        Write-Host "`nUsage Examples:" -ForegroundColor Yellow
        Write-Host "  Test CSV: .\PurchaseTracking.ps1 -InputCSV 'file.csv' -TestCSV"
        Write-Host "  Import: .\PurchaseTracking.ps1 -InputCSV 'file.csv' -DatabasePath 'data.db'"
        Write-Host "  Update: .\PurchaseTracking.ps1 -DatabasePath 'data.db' -UpdateExisting"
        Write-Host "  Import with price lookup: .\PurchaseTracking.ps1 -InputCSV 'file.csv' -UseAPI"
        exit 0
    }

    # Initialize database if needed
    if (-not (Test-Path $DatabasePath)) {
        Write-Host "Creating new database..." -ForegroundColor Yellow
        Initialize-PurchaseDatabase -DatabasePath $DatabasePath
    }

    # Process data using module function
    Write-Host "`n=== Starting Item Processing ===" -ForegroundColor Cyan

    $progressCallback = {
        param($message)
        Write-Host $message
    }

    Invoke-ItemProcessing `
        -DatabasePath $DatabasePath `
        -CSVPath $InputCSV `
        -UpdateExisting:$UpdateExisting `
        -PerformPriceLookup:$UseAPI `
        -ProgressCallback $progressCallback

    # Generate summary
    New-SummaryReport -DatabasePath $DatabasePath

    if ($ShowDashboard) {
        Write-Host "`nNote: Dashboard functionality requires the GUI script (PurchaseTrackerGUI.ps1)" -ForegroundColor Yellow
    }

    Write-Host "`n=== Processing Complete ===" -ForegroundColor Green
    Write-Host "Database location: $DatabasePath" -ForegroundColor Gray
}
catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}

#endregion