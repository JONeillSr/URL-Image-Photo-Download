<#
.SYNOPSIS
    Enhanced GUI Application for Purchase Tracking with Price Lookup and Visual Analytics.

.DESCRIPTION
    A complete Windows Forms GUI application for managing purchase tracking data with
    inline editing, automated online price lookup, visual charts, and real-time saving.
    Uses shared module for business logic.

.PARAMETER DatabasePath
    Path to the SQLite database file. Creates new if doesn't exist.

.EXAMPLE
    .\PurchaseTrackerGUI.ps1
    .\PurchaseTrackerGUI.ps1 -DatabasePath "C:\Data\purchases.db"

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10/10/2025
    Version: 2.9.0
    Change Date: 10/13/2025
    Change Purpose: Enhanced copy feature to include all product details, fixed confirmation dialog display

.CHANGELOG
    2.9.0 - Fixed copy feature to include Brand, Model, PartNumber, Color, and MSRP; fixed confirmation dialog to show actual values
    2.8.0 - Added Brand and Model columns to datagrid with inline editing, enhanced search to include brand and model
    2.7.0 - Added PartNumber, Color, and MSRP columns to datagrid with inline editing, enhanced search to include part numbers
    2.6.0 - Added "Copy Prices to Matching Descriptions" feature, fixed null value handling in grid updates
    2.2.0 - Refactored to use PurchaseTracking.psm1 module, added batch processing dialog
    2.1.0 - Complete rewrite of CellEndEdit handler, fixed all recalculation issues
    2.0.6 - Added type conversion for numeric fields and detailed status messages
    2.0.5 - Added Refresh button to Items tab next to Search box
    2.0.4 - Fixed recalculation by updating grid cells directly with InvalidateRow
    2.0.3 - Fixed profit/margin recalculation by updating underlying DataTable
    2.0.2 - Changed to CellEndEdit event, fixed real-time profit/margin recalculation
    2.0.1 - Fixed $sender variable warning, added auto-recalculation of profit/margin on edits
    2.0.0 - Added inline grid editing, online price lookup, visual charts, real-time save
    1.2.0 - Fixed menu separator, removed ShouldProcess from GUI functions, cleaned up code
    1.1.0 - Fixed PSScriptAnalyzer warnings
    1.0.0 - Initial GUI release
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$DatabasePath = ".\PurchaseTracking.db"
)

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Module Import

# Import shared module
$modulePath = Join-Path $PSScriptRoot "modules\PurchaseTracking.psm1"
if (Test-Path $modulePath) {
    try {
        Import-Module $modulePath -Force -ErrorAction Stop
        Write-Verbose "Successfully loaded PurchaseTracking module"
        $script:ModuleLoaded = $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load PurchaseTracking module: $_`n`nBatch processing features will be disabled.",
            "Module Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $script:ModuleLoaded = $false
    }
} else {
    [System.Windows.Forms.MessageBox]::Show(
        "PurchaseTracking.psm1 module not found in script directory.`n`nBatch processing features will be disabled.",
        "Module Missing",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    $script:ModuleLoaded = $false
}

# Import SQLite module
try {
    Import-Module PSSQLite -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "PSSQLite module not found. Please install: Install-Module PSSQLite -Scope CurrentUser",
        "Module Missing",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

#endregion

# Global Variables
$script:DatabasePath = $DatabasePath
$script:AutoSaveEnabled = $true

#region Database Functions (Local wrappers for module functions)

function Initialize-Database {
    if (-not (Test-Path $script:DatabasePath)) {
        if ($script:ModuleLoaded) {
            Initialize-PurchaseDatabase -DatabasePath $script:DatabasePath
        }
        else {
            # Fallback: Create basic schema
            $schema = @"
CREATE TABLE IF NOT EXISTS Items (
    ItemID INTEGER PRIMARY KEY AUTOINCREMENT,
    Lot INTEGER,
    Description TEXT NOT NULL,
    Category TEXT,
    Quantity INTEGER DEFAULT 1,
    PerItemCost REAL,
    TotalCost REAL,
    CurrentMSRP REAL,
    CurrentMarketAvg REAL,
    SuggestedPrice REAL,
    MinAcceptablePrice REAL,
    SoldPrice REAL,
    Notes TEXT,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    IsActive BOOLEAN DEFAULT 1
);

CREATE TABLE IF NOT EXISTS PriceHistory (
    PriceID INTEGER PRIMARY KEY AUTOINCREMENT,
    ItemID INTEGER,
    Source TEXT,
    Price REAL,
    Confidence REAL,
    CaptureDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ItemID) REFERENCES Items(ItemID)
);
"@
            Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $schema
        }
    }
}

function Get-PurchaseItemLocal {
    param([string]$Filter = "")

    if ($script:ModuleLoaded) {
        return Get-PurchaseItem -DatabasePath $script:DatabasePath -Filter $Filter
    }
    else {
        # Fallback query
        $query = @"
            SELECT
                ItemID,
                Lot,
                Description,
                Brand,
                Model,
                PartNumber,
                Color,
                Category,
                Quantity,
                ROUND(PerItemCost, 2) as PerItemCost,
                ROUND(CurrentMSRP, 2) as MSRP,
                ROUND(CurrentMarketAvg, 2) as CompetitorPrice,
                ROUND(SuggestedPrice, 2) as OurPrice,
                ROUND((SuggestedPrice - PerItemCost), 2) as PerItemProfit,
                CASE WHEN PerItemCost > 0
                    THEN ROUND((SuggestedPrice/PerItemCost - 1) * 100, 1)
                    ELSE 0 END as Margin
            FROM Items
            WHERE IsActive = 1
            $(if ($Filter) { "AND (Description LIKE '%$Filter%' OR PartNumber LIKE '%$Filter%' OR Brand LIKE '%$Filter%' OR Model LIKE '%$Filter%')" })
            ORDER BY ItemID
"@
        return Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query
    }
}

function Update-ItemField {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [int]$ItemID,
        [string]$FieldName,
        [object]$Value
    )

    if ($script:ModuleLoaded) {
        return Update-PurchaseItemField -DatabasePath $script:DatabasePath -ItemID $ItemID -FieldName $FieldName -Value $Value
    }
    else {
        # Fallback
        try {
            $query = "UPDATE Items SET $FieldName = @Value, ModifiedDate = CURRENT_TIMESTAMP WHERE ItemID = @ID"
            $params = @{ Value = $Value; ID = $ItemID }
            Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query -SqlParameters $params
            return $true
        }
        catch {
            return $false
        }
    }
}

#endregion

#region Import Summary Dialog

function Show-ImportSummary {
    param(
        [int]$TotalRows,
        [int]$Added,
        [int]$Updated,
        [int]$Skipped,
        [int]$Errors,
        [string[]]$ErrorMessages = @()
    )

    $summaryForm = New-Object System.Windows.Forms.Form
    $summaryForm.Text = "Import Summary"
    $summaryForm.Size = New-Object System.Drawing.Size(500, 450)
    $summaryForm.StartPosition = "CenterParent"
    $summaryForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $summaryForm.MaximizeBox = $false
    $summaryForm.BackColor = [System.Drawing.Color]::White

    # Title
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Import Complete"
    $lblTitle.Location = New-Object System.Drawing.Point(20, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(450, 30)
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(33, 150, 243)

    # Statistics Panel
    $statsPanel = New-Object System.Windows.Forms.Panel
    $statsPanel.Location = New-Object System.Drawing.Point(20, 60)
    $statsPanel.Size = New-Object System.Drawing.Size(450, 200)
    $statsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statsPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    $yPos = 15

    # Total Rows
    $lblTotal = New-Object System.Windows.Forms.Label
    $lblTotal.Text = "Total Rows in CSV:"
    $lblTotal.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblTotal.Size = New-Object System.Drawing.Size(200, 25)
    $lblTotal.Font = New-Object System.Drawing.Font("Segoe UI", 11)

    $lblTotalValue = New-Object System.Windows.Forms.Label
    $lblTotalValue.Text = $TotalRows
    $lblTotalValue.Location = New-Object System.Drawing.Point(350, $yPos)
    $lblTotalValue.Size = New-Object System.Drawing.Size(80, 25)
    $lblTotalValue.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblTotalValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

    $yPos += 35

    # Separator
    $separator1 = New-Object System.Windows.Forms.Label
    $separator1.Location = New-Object System.Drawing.Point(20, $yPos)
    $separator1.Size = New-Object System.Drawing.Size(410, 2)
    $separator1.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

    $yPos += 15

    # Added
    $lblAdded = New-Object System.Windows.Forms.Label
    $lblAdded.Text = "✓ New Items Added:"
    $lblAdded.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblAdded.Size = New-Object System.Drawing.Size(200, 25)
    $lblAdded.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $lblAdded.ForeColor = [System.Drawing.Color]::FromArgb(76, 175, 80)

    $lblAddedValue = New-Object System.Windows.Forms.Label
    $lblAddedValue.Text = $Added
    $lblAddedValue.Location = New-Object System.Drawing.Point(350, $yPos)
    $lblAddedValue.Size = New-Object System.Drawing.Size(80, 25)
    $lblAddedValue.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblAddedValue.ForeColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $lblAddedValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

    $yPos += 30

    # Updated
    $lblUpdated = New-Object System.Windows.Forms.Label
    $lblUpdated.Text = "↻ Existing Items Updated:"
    $lblUpdated.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblUpdated.Size = New-Object System.Drawing.Size(200, 25)
    $lblUpdated.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $lblUpdated.ForeColor = [System.Drawing.Color]::FromArgb(33, 150, 243)

    $lblUpdatedValue = New-Object System.Windows.Forms.Label
    $lblUpdatedValue.Text = $Updated
    $lblUpdatedValue.Location = New-Object System.Drawing.Point(350, $yPos)
    $lblUpdatedValue.Size = New-Object System.Drawing.Size(80, 25)
    $lblUpdatedValue.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblUpdatedValue.ForeColor = [System.Drawing.Color]::FromArgb(33, 150, 243)
    $lblUpdatedValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

    $yPos += 30

    # Skipped
    if ($Skipped -gt 0) {
        $lblSkipped = New-Object System.Windows.Forms.Label
        $lblSkipped.Text = "⊘ Items Skipped:"
        $lblSkipped.Location = New-Object System.Drawing.Point(20, $yPos)
        $lblSkipped.Size = New-Object System.Drawing.Size(200, 25)
        $lblSkipped.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $lblSkipped.ForeColor = [System.Drawing.Color]::FromArgb(255, 152, 0)

        $lblSkippedValue = New-Object System.Windows.Forms.Label
        $lblSkippedValue.Text = $Skipped
        $lblSkippedValue.Location = New-Object System.Drawing.Point(350, $yPos)
        $lblSkippedValue.Size = New-Object System.Drawing.Size(80, 25)
        $lblSkippedValue.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $lblSkippedValue.ForeColor = [System.Drawing.Color]::FromArgb(255, 152, 0)
        $lblSkippedValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

        $statsPanel.Controls.AddRange(@($lblSkipped, $lblSkippedValue))
        $yPos += 30
    }

    # Errors
    if ($Errors -gt 0) {
        $lblErrors = New-Object System.Windows.Forms.Label
        $lblErrors.Text = "✗ Errors:"
        $lblErrors.Location = New-Object System.Drawing.Point(20, $yPos)
        $lblErrors.Size = New-Object System.Drawing.Size(200, 25)
        $lblErrors.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $lblErrors.ForeColor = [System.Drawing.Color]::FromArgb(244, 67, 54)

        $lblErrorsValue = New-Object System.Windows.Forms.Label
        $lblErrorsValue.Text = $Errors
        $lblErrorsValue.Location = New-Object System.Drawing.Point(350, $yPos)
        $lblErrorsValue.Size = New-Object System.Drawing.Size(80, 25)
        $lblErrorsValue.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $lblErrorsValue.ForeColor = [System.Drawing.Color]::FromArgb(244, 67, 54)
        $lblErrorsValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

        $statsPanel.Controls.AddRange(@($lblErrors, $lblErrorsValue))
    }

    $statsPanel.Controls.AddRange(@(
        $lblTotal, $lblTotalValue, $separator1,
        $lblAdded, $lblAddedValue,
        $lblUpdated, $lblUpdatedValue
    ))

    # Error Details (if any)
    if ($Errors -gt 0 -and $ErrorMessages.Count -gt 0) {
        $lblErrorsTitle = New-Object System.Windows.Forms.Label
        $lblErrorsTitle.Text = "Error Details:"
        $lblErrorsTitle.Location = New-Object System.Drawing.Point(20, 270)
        $lblErrorsTitle.Size = New-Object System.Drawing.Size(450, 20)
        $lblErrorsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lblErrorsTitle.ForeColor = [System.Drawing.Color]::FromArgb(244, 67, 54)

        $txtErrors = New-Object System.Windows.Forms.TextBox
        $txtErrors.Location = New-Object System.Drawing.Point(20, 295)
        $txtErrors.Size = New-Object System.Drawing.Size(450, 70)
        $txtErrors.Multiline = $true
        $txtErrors.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $txtErrors.ReadOnly = $true
        $txtErrors.Font = New-Object System.Drawing.Font("Consolas", 9)
        $txtErrors.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 245)
        $txtErrors.Text = ($ErrorMessages | Select-Object -First 5) -join "`r`n"

        $summaryForm.Controls.AddRange(@($lblErrorsTitle, $txtErrors))
    }

    # Success Message
    $successMsg = New-Object System.Windows.Forms.Label
    $successMsgY = if ($Errors -gt 0) { 375 } else { 270 }
    $successMsg.Location = New-Object System.Drawing.Point -ArgumentList 20, $successMsgY
    $successMsg.Size = New-Object System.Drawing.Size(450, 30)
    $successMsg.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $successMsg.ForeColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $successMsg.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    if ($Errors -eq 0) {
        $successMsg.Text = "✓ All items imported successfully!"
    } else {
        $successMsg.Text = "⚠ Import completed with some errors"
        $successMsg.ForeColor = [System.Drawing.Color]::FromArgb(255, 152, 0)
    }

    # Close Button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnCloseY = if ($Errors -gt 0) { 410 } else { 310 }
    $btnClose.Location = New-Object System.Drawing.Point -ArgumentList 190, $btnCloseY
    $btnClose.Size = New-Object System.Drawing.Size(120, 30)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(33, 150, 243)
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnClose.Add_Click({ $summaryForm.Close() })

    $summaryForm.Controls.AddRange(@($lblTitle, $statsPanel, $successMsg, $btnClose))

    [void]$summaryForm.ShowDialog()
}

#endregion

#region Batch Processing Functions

function Show-BatchProcessDialog {
    if (-not $script:ModuleLoaded) {
        [System.Windows.Forms.MessageBox]::Show(
            "Batch processing requires the PurchaseTracking.psm1 module.`n`nPlease ensure the module is in the same directory as this script.",
            "Feature Unavailable",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Batch Process Items"
    $dialog.Size = New-Object System.Drawing.Size(500, 550)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false

    # CSV File Selection
    $lblCSV = New-Object System.Windows.Forms.Label
    $lblCSV.Text = "CSV File:"
    $lblCSV.Location = New-Object System.Drawing.Point(10, 20)
    $lblCSV.Size = New-Object System.Drawing.Size(80, 20)

    $txtCSV = New-Object System.Windows.Forms.TextBox
    $txtCSV.Location = New-Object System.Drawing.Point(100, 20)
    $txtCSV.Size = New-Object System.Drawing.Size(300, 20)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = New-Object System.Drawing.Point(410, 18)
    $btnBrowse.Size = New-Object System.Drawing.Size(70, 23)
    $btnBrowse.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtCSV.Text = $openFileDialog.FileName
        }
    })

    # Options
    $chkUpdateExisting = New-Object System.Windows.Forms.CheckBox
    $chkUpdateExisting.Text = "Update existing items with new prices"
    $chkUpdateExisting.Location = New-Object System.Drawing.Point(20, 60)
    $chkUpdateExisting.Size = New-Object System.Drawing.Size(400, 20)

    $chkPriceLookup = New-Object System.Windows.Forms.CheckBox
    $chkPriceLookup.Text = "Perform online price lookup"
    $chkPriceLookup.Location = New-Object System.Drawing.Point(20, 90)
    $chkPriceLookup.Size = New-Object System.Drawing.Size(400, 20)
    $chkPriceLookup.Checked = $true

    # Sites to search
    $grpSites = New-Object System.Windows.Forms.GroupBox
    $grpSites.Text = "Price Lookup Sites"
    $grpSites.Location = New-Object System.Drawing.Point(20, 120)
    $grpSites.Size = New-Object System.Drawing.Size(450, 180)

    # RV-Specific Sites (Left Column)
    $lblRV = New-Object System.Windows.Forms.Label
    $lblRV.Text = "RV/Trailer Sites:"
    $lblRV.Location = New-Object System.Drawing.Point(10, 20)
    $lblRV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $chkLippert = New-Object System.Windows.Forms.CheckBox
    $chkLippert.Text = "Lippert"
    $chkLippert.Location = New-Object System.Drawing.Point(10, 40)
    $chkLippert.Checked = $true

    $chkCampingWorld = New-Object System.Windows.Forms.CheckBox
    $chkCampingWorld.Text = "Camping World"
    $chkCampingWorld.Location = New-Object System.Drawing.Point(10, 65)
    $chkCampingWorld.Checked = $true

    $chkUnitedRV = New-Object System.Windows.Forms.CheckBox
    $chkUnitedRV.Text = "United RV"
    $chkUnitedRV.Location = New-Object System.Drawing.Point(10, 90)
    $chkUnitedRV.Checked = $true

    $chkEtrailer = New-Object System.Windows.Forms.CheckBox
    $chkEtrailer.Text = "eTrailer"
    $chkEtrailer.Location = New-Object System.Drawing.Point(10, 115)
    $chkEtrailer.Checked = $true

    $chkRVPartsCountry = New-Object System.Windows.Forms.CheckBox
    $chkRVPartsCountry.Text = "RV Parts Country"
    $chkRVPartsCountry.Location = New-Object System.Drawing.Point(10, 140)

    # General Sites (Right Column)
    $lblGeneral = New-Object System.Windows.Forms.Label
    $lblGeneral.Text = "General Sites:"
    $lblGeneral.Location = New-Object System.Drawing.Point(230, 20)
    $lblGeneral.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $chkAmazon = New-Object System.Windows.Forms.CheckBox
    $chkAmazon.Text = "Amazon"
    $chkAmazon.Location = New-Object System.Drawing.Point(230, 40)
    $chkAmazon.Checked = $true

    $chkEbay = New-Object System.Windows.Forms.CheckBox
    $chkEbay.Text = "eBay"
    $chkEbay.Location = New-Object System.Drawing.Point(230, 65)
    $chkEbay.Checked = $true

    $chkWalmart = New-Object System.Windows.Forms.CheckBox
    $chkWalmart.Text = "Walmart"
    $chkWalmart.Location = New-Object System.Drawing.Point(230, 90)

    $grpSites.Controls.AddRange(@(
        $lblRV, $chkLippert, $chkCampingWorld, $chkUnitedRV, $chkEtrailer, $chkRVPartsCountry,
        $lblGeneral, $chkAmazon, $chkEbay, $chkWalmart
    ))

    # Progress display
    $txtProgress = New-Object System.Windows.Forms.TextBox
    $txtProgress.Location = New-Object System.Drawing.Point(20, 310)
    $txtProgress.Size = New-Object System.Drawing.Size(450, 140)
    $txtProgress.Multiline = $true
    $txtProgress.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtProgress.ReadOnly = $true
    $txtProgress.Font = New-Object System.Drawing.Font("Consolas", 9)

    # Buttons
    $btnProcess = New-Object System.Windows.Forms.Button
    $btnProcess.Text = "Start Processing"
    $btnProcess.Location = New-Object System.Drawing.Point(250, 465)
    $btnProcess.Size = New-Object System.Drawing.Size(110, 30)
    $btnProcess.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnProcess.ForeColor = [System.Drawing.Color]::White
    $btnProcess.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnProcess.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtCSV.Text)) {
            if (-not $chkUpdateExisting.Checked) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please select a CSV file or check 'Update existing items'",
                    "Missing Input",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
        }

        # Build sites list
        $sites = @()
        if ($chkLippert.Checked) { $sites += 'lippert.com' }
        if ($chkCampingWorld.Checked) { $sites += 'campingworld.com' }
        if ($chkUnitedRV.Checked) { $sites += 'unitedrv.com' }
        if ($chkEtrailer.Checked) { $sites += 'etrailer.com' }
        if ($chkRVPartsCountry.Checked) { $sites += 'rvpartscountry.com' }
        if ($chkAmazon.Checked) { $sites += 'amazon.com' }
        if ($chkEbay.Checked) { $sites += 'ebay.com' }
        if ($chkWalmart.Checked) { $sites += 'walmart.com' }

        # Disable button during processing
        $btnProcess.Enabled = $false
        $txtProgress.Text = "Starting batch processing...`r`n"
        $dialog.Refresh()

        # Track statistics
        $script:ImportStats = @{
            TotalRows = 0
            Added = 0
            Updated = 0
            Skipped = 0
            Errors = 0
            ErrorMessages = @()
        }

        # Count total rows if CSV provided
        if ($txtCSV.Text -and (Test-Path $txtCSV.Text)) {
            $csvData = Import-Csv -Path $txtCSV.Text
            $script:ImportStats.TotalRows = $csvData.Count
            $txtProgress.AppendText("CSV contains $($csvData.Count) rows`r`n`r`n")
        }

        try {
            $progressCallback = {
                param($message)

                # Parse statistics from messages
                if ($message -match 'Added to database') {
                    $script:ImportStats.Added++
                }
                elseif ($message -match 'Item exists.*Updating') {
                    $script:ImportStats.Updated++
                }
                elseif ($message -match 'ERROR:') {
                    $script:ImportStats.Errors++
                    $script:ImportStats.ErrorMessages += $message
                }

                $txtProgress.AppendText("$message`r`n")
                $txtProgress.SelectionStart = $txtProgress.Text.Length
                $txtProgress.ScrollToCaret()
                $txtProgress.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
            }

            Invoke-ItemProcessing `
                -DatabasePath $script:DatabasePath `
                -CSVPath $txtCSV.Text `
                -UpdateExisting:$chkUpdateExisting.Checked `
                -PerformPriceLookup:$chkPriceLookup.Checked `
                -PriceLookupSites $sites `
                -ProgressCallback $progressCallback

            $txtProgress.AppendText("`r`n=== Processing Complete ===`r`n")

            # Refresh all displays
            Update-Dashboard
            Update-ItemGrid
            Update-Charts

            # Show import summary
            Show-ImportSummary `
                -TotalRows $script:ImportStats.TotalRows `
                -Added $script:ImportStats.Added `
                -Updated $script:ImportStats.Updated `
                -Skipped $script:ImportStats.Skipped `
                -Errors $script:ImportStats.Errors `
                -ErrorMessages $script:ImportStats.ErrorMessages
        }
        catch {
            $txtProgress.AppendText("`r`nERROR: $_`r`n")
            [System.Windows.Forms.MessageBox]::Show(
                "Error during processing: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        finally {
            $btnProcess.Enabled = $true
        }
    })

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(370, 465)
    $btnClose.Size = New-Object System.Drawing.Size(100, 30)
    $btnClose.Add_Click({ $dialog.Close() })

    $dialog.Controls.AddRange(@(
        $lblCSV, $txtCSV, $btnBrowse,
        $chkUpdateExisting, $chkPriceLookup,
        $grpSites, $txtProgress,
        $btnProcess, $btnClose
    ))

    [void]$dialog.ShowDialog()
}

#endregion

#region Price Lookup Functions

function Search-OnlinePrice {
    param(
        [string]$ItemDescription,
        [int]$ItemID
    )

    # Create progress form
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Searching Prices..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Searching for: $ItemDescription"
    $progressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 40)
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 70)
    $progressBar.Size = New-Object System.Drawing.Size(350, 23)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee

    $progressForm.Controls.AddRange(@($progressLabel, $progressBar))
    $progressForm.Show()
    $progressForm.Refresh()

    try {
        if ($script:ModuleLoaded) {
            # Use module function
            $result = Invoke-PriceResearch `
                -DatabasePath $script:DatabasePath `
                -ItemID $ItemID `
                -Description $ItemDescription `
                -Sites @('lippert.com', 'campingworld.com', 'unitedrv.com', 'etrailer.com', 'amazon.com', 'ebay.com')

            $progressForm.Close()

            if ($result.Success) {
                $resultMsg = "Market Price Research Results:`n`n"
                $resultMsg += "Average Market Price: `$$($result.AveragePrice)`n"
                $resultMsg += "MSRP Price: `$$($result.MSRPPrice)`n"
                $resultMsg += "Sources found: $($result.PriceCount)"

                [System.Windows.Forms.MessageBox]::Show(
                    $resultMsg,
                    "Price Lookup Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

                return $result.AveragePrice
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "No pricing data found for this item.",
                    "Price Lookup",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                return $null
            }
        }
        else {
            # Fallback: Simulate search
            Start-Sleep -Milliseconds 1500
            $progressForm.Close()

            [System.Windows.Forms.MessageBox]::Show(
                "Price lookup requires the PurchaseTracking.psm1 module.",
                "Feature Unavailable",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return $null
        }
    }
    catch {
        $progressForm.Close()
        [System.Windows.Forms.MessageBox]::Show(
            "Error during price lookup: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $null
    }
}

#endregion

#region Main Form

# Create the main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Purchase Tracker Pro v2.9.0"
$mainForm.Size = New-Object System.Drawing.Size(1200, 800)
$mainForm.StartPosition = "CenterScreen"

# Create MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$menuStrip.ForeColor = [System.Drawing.Color]::White

# File Menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

$batchProcessMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$batchProcessMenuItem.Text = "Batch Process with Price Lookup"
$batchProcessMenuItem.Add_Click({ Show-BatchProcessDialog })

$importMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$importMenuItem.Text = "Import CSV"
$importMenuItem.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Import-CSVData -Path $openFileDialog.FileName
    }
})

$exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportMenuItem.Text = "Export CSV"
$exportMenuItem.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.FileName = "Export_$(Get-Date -Format 'yyyyMMdd').csv"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-CSVData -Path $saveFileDialog.FileName
    }
})

$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.Add_Click({ $mainForm.Close() })

[void]$fileMenu.DropDownItems.Add($batchProcessMenuItem)
[void]$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$fileMenu.DropDownItems.Add($importMenuItem)
[void]$fileMenu.DropDownItems.Add($exportMenuItem)
[void]$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$fileMenu.DropDownItems.Add($exitMenuItem)

[void]$menuStrip.Items.Add($fileMenu)

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 30)
$tabControl.Size = New-Object System.Drawing.Size(1160, 700)

# Dashboard Tab
$dashboardTab = New-Object System.Windows.Forms.TabPage
$dashboardTab.Text = "Dashboard"
$dashboardTab.BackColor = [System.Drawing.Color]::WhiteSmoke

# Stats Panel
$statsPanel = New-Object System.Windows.Forms.Panel
$statsPanel.Location = New-Object System.Drawing.Point(10, 10)
$statsPanel.Size = New-Object System.Drawing.Size(1140, 120)
$statsPanel.BackColor = [System.Drawing.Color]::White
$statsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

# Stats Labels
$lblTotalItems = New-Object System.Windows.Forms.Label
$lblTotalItems.Text = "Total Items:"
$lblTotalItems.Location = New-Object System.Drawing.Point(20, 20)
$lblTotalItems.Size = New-Object System.Drawing.Size(100, 20)
$lblTotalItems.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lblTotalItemsValue = New-Object System.Windows.Forms.Label
$lblTotalItemsValue.Text = "0"
$lblTotalItemsValue.Location = New-Object System.Drawing.Point(20, 45)
$lblTotalItemsValue.Size = New-Object System.Drawing.Size(100, 30)
$lblTotalItemsValue.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTotalItemsValue.ForeColor = [System.Drawing.Color]::FromArgb(0, 122, 204)

$lblInvestment = New-Object System.Windows.Forms.Label
$lblInvestment.Text = "Total Investment:"
$lblInvestment.Location = New-Object System.Drawing.Point(250, 20)
$lblInvestment.Size = New-Object System.Drawing.Size(150, 20)
$lblInvestment.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lblInvestmentValue = New-Object System.Windows.Forms.Label
$lblInvestmentValue.Text = "$0.00"
$lblInvestmentValue.Location = New-Object System.Drawing.Point(250, 45)
$lblInvestmentValue.Size = New-Object System.Drawing.Size(150, 30)
$lblInvestmentValue.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblInvestmentValue.ForeColor = [System.Drawing.Color]::FromArgb(244, 67, 54)

$lblRevenue = New-Object System.Windows.Forms.Label
$lblRevenue.Text = "Potential Revenue:"
$lblRevenue.Location = New-Object System.Drawing.Point(500, 20)
$lblRevenue.Size = New-Object System.Drawing.Size(150, 20)
$lblRevenue.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lblRevenueValue = New-Object System.Windows.Forms.Label
$lblRevenueValue.Text = "$0.00"
$lblRevenueValue.Location = New-Object System.Drawing.Point(500, 45)
$lblRevenueValue.Size = New-Object System.Drawing.Size(150, 30)
$lblRevenueValue.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblRevenueValue.ForeColor = [System.Drawing.Color]::FromArgb(76, 175, 80)

$lblMargin = New-Object System.Windows.Forms.Label
$lblMargin.Text = "Average Margin:"
$lblMargin.Location = New-Object System.Drawing.Point(750, 20)
$lblMargin.Size = New-Object System.Drawing.Size(150, 20)
$lblMargin.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lblMarginValue = New-Object System.Windows.Forms.Label
$lblMarginValue.Text = "0%"
$lblMarginValue.Location = New-Object System.Drawing.Point(750, 45)
$lblMarginValue.Size = New-Object System.Drawing.Size(150, 30)
$lblMarginValue.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblMarginValue.ForeColor = [System.Drawing.Color]::FromArgb(76, 175, 80)

$statsPanel.Controls.AddRange(@(
    $lblTotalItems, $lblTotalItemsValue,
    $lblInvestment, $lblInvestmentValue,
    $lblRevenue, $lblRevenueValue,
    $lblMargin, $lblMarginValue
))

# Action Buttons Panel
$actionsPanel = New-Object System.Windows.Forms.Panel
$actionsPanel.Location = New-Object System.Drawing.Point(10, 140)
$actionsPanel.Size = New-Object System.Drawing.Size(1140, 60)
$actionsPanel.BackColor = [System.Drawing.Color]::White

$btnAddItem = New-Object System.Windows.Forms.Button
$btnAddItem.Text = "Add Item"
$btnAddItem.Location = New-Object System.Drawing.Point(20, 15)
$btnAddItem.Size = New-Object System.Drawing.Size(100, 30)
$btnAddItem.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$btnAddItem.ForeColor = [System.Drawing.Color]::White
$btnAddItem.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnAddItem.Add_Click({ Show-AddItemDialog })

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Location = New-Object System.Drawing.Point(130, 15)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 30)
$btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$btnRefresh.ForeColor = [System.Drawing.Color]::White
$btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefresh.Add_Click({ Update-Dashboard; Update-ItemGrid; Update-Charts })

$actionsPanel.Controls.AddRange(@($btnAddItem, $btnRefresh))

$dashboardTab.Controls.AddRange(@($statsPanel, $actionsPanel))

# Items Tab
$itemsTab = New-Object System.Windows.Forms.TabPage
$itemsTab.Text = "Items"

# Search Panel
$searchPanel = New-Object System.Windows.Forms.Panel
$searchPanel.Location = New-Object System.Drawing.Point(10, 10)
$searchPanel.Size = New-Object System.Drawing.Size(1140, 40)
$searchPanel.BackColor = [System.Drawing.Color]::White

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.Location = New-Object System.Drawing.Point(10, 10)
$lblSearch.Size = New-Object System.Drawing.Size(60, 20)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(75, 8)
$txtSearch.Size = New-Object System.Drawing.Size(300, 20)
$txtSearch.Add_TextChanged({
    Update-ItemGrid -Filter $txtSearch.Text
})

$btnRefreshItems = New-Object System.Windows.Forms.Button
$btnRefreshItems.Text = "Refresh"
$btnRefreshItems.Location = New-Object System.Drawing.Point(385, 7)
$btnRefreshItems.Size = New-Object System.Drawing.Size(80, 23)
$btnRefreshItems.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$btnRefreshItems.ForeColor = [System.Drawing.Color]::White
$btnRefreshItems.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefreshItems.Add_Click({ Update-ItemGrid -Filter $txtSearch.Text; $statusLabel.Text = "Grid refreshed" })

$searchPanel.Controls.AddRange(@($lblSearch, $txtSearch, $btnRefreshItems))

# Items Grid with inline editing
$itemsGrid = New-Object System.Windows.Forms.DataGridView
$itemsGrid.Location = New-Object System.Drawing.Point(10, 60)
$itemsGrid.Size = New-Object System.Drawing.Size(1140, 600)
$itemsGrid.AllowUserToAddRows = $false
$itemsGrid.AllowUserToDeleteRows = $false
$itemsGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$itemsGrid.MultiSelect = $false
$itemsGrid.BackgroundColor = [System.Drawing.Color]::White
$itemsGrid.ReadOnly = $false
$itemsGrid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2

# Add context menu for price lookup and delete
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$lookupPriceItem = New-Object System.Windows.Forms.ToolStripMenuItem
$lookupPriceItem.Text = "Search Online Price"
$lookupPriceItem.Add_Click({
    if ($itemsGrid.SelectedRows.Count -gt 0) {
        $row = $itemsGrid.SelectedRows[0]
        $itemID = $row.Cells["ID"].Value
        $description = $row.Cells["Description"].Value

        $avgPrice = Search-OnlinePrice -ItemDescription $description -ItemID $itemID

        if ($avgPrice) {
            Update-Dashboard
            Update-ItemGrid
            Update-Charts
        }
    }
})

$deleteItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteItem.Text = "Delete Item"
$deleteItem.Add_Click({
    if ($itemsGrid.SelectedRows.Count -gt 0) {
        $row = $itemsGrid.SelectedRows[0]
        $itemID = $row.Cells["ID"].Value
        $description = $row.Cells["Description"].Value

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete this item?`n`n$description",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                # Soft delete by setting IsActive to 0
                $query = "UPDATE Items SET IsActive = 0, ModifiedDate = CURRENT_TIMESTAMP WHERE ItemID = @ID"
                $params = @{ ID = $itemID }
                Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query -SqlParameters $params

                Update-Dashboard
                Update-ItemGrid
                Update-Charts

                $statusLabel.Text = "Item #$itemID deleted successfully"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error deleting item: $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    }
})

$copyPricesItem = New-Object System.Windows.Forms.ToolStripMenuItem
$copyPricesItem.Text = "Copy Product Details to Matching Descriptions"
$copyPricesItem.Add_Click({
    if ($itemsGrid.SelectedRows.Count -gt 0) {
        $row = $itemsGrid.SelectedRows[0]
        $itemID = $row.Cells["ID"].Value
        $description = $row.Cells["Description"].Value

        # Get all values from the row
        $brandValue = $row.Cells["Brand"].Value
        $modelValue = $row.Cells["Model"].Value
        $partNumberValue = $row.Cells["PartNumber"].Value
        $colorValue = $row.Cells["Color"].Value
        $msrpValue = $row.Cells["MSRP"].Value
        $competitorPriceValue = $row.Cells["CompetitorPrice"].Value
        $ourPriceValue = $row.Cells["Our Price"].Value

        if ($competitorPriceValue -eq 0 -and $ourPriceValue -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "This item has no prices to copy.",
                "No Prices",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        # Find matching items
        $findQuery = @"
            SELECT ItemID, Lot
            FROM Items
            WHERE Description = @Description
            AND ItemID != @ItemID
            AND IsActive = 1
"@
        $findParams = @{ Description = $description; ItemID = $itemID }
        $matchingItems = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $findQuery -SqlParameters $findParams

        if ($matchingItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No other items found with matching description.",
                "No Matches",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        # Build confirmation message
        $lotList = ($matchingItems | ForEach-Object { "Lot $($_.Lot)" }) -join ", "
        $compPriceDisplay = [Math]::Round($competitorPriceValue, 2)
        $ourPriceDisplay = [Math]::Round($ourPriceValue, 2)
        $msrpDisplay = [Math]::Round($msrpValue, 2)

        $confirmMessage = "Copy product details from this item to $($matchingItems.Count) other item(s)?`n`n"
        $confirmMessage += "Matching lots: $lotList`n`n"
        $confirmMessage += "Brand: $brandValue`n"
        $confirmMessage += "Model: $modelValue`n"
        $confirmMessage += "Part Number: $partNumberValue`n"
        $confirmMessage += "Color: $colorValue`n"
        $confirmMessage += "MSRP: `$$msrpDisplay`n"
        $confirmMessage += "Competitor Price: `$$compPriceDisplay`n"
        $confirmMessage += "Our Price: `$$ourPriceDisplay"

        $result = [System.Windows.Forms.MessageBox]::Show(
            $confirmMessage,
            "Confirm Copy Product Details",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $updateQuery = @"
                    UPDATE Items
                    SET Brand = @Brand,
                        Model = @Model,
                        PartNumber = @PartNumber,
                        Color = @Color,
                        CurrentMSRP = @MSRP,
                        CurrentMarketAvg = @CompPrice,
                        SuggestedPrice = @OurPrice,
                        ModifiedDate = CURRENT_TIMESTAMP
                    WHERE Description = @Description
                    AND ItemID != @ItemID
                    AND IsActive = 1
"@
                $updateParams = @{
                    Brand = $brandValue
                    Model = $modelValue
                    PartNumber = $partNumberValue
                    Color = $colorValue
                    MSRP = $msrpValue
                    CompPrice = $competitorPriceValue
                    OurPrice = $ourPriceValue
                    Description = $description
                    ItemID = $itemID
                }

                Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $updateQuery -SqlParameters $updateParams

                Update-Dashboard
                Update-ItemGrid
                Update-Charts

                $statusLabel.Text = "Copied product details to $($matchingItems.Count) matching item(s)"

                [System.Windows.Forms.MessageBox]::Show(
                    "Successfully copied product details to $($matchingItems.Count) item(s)!",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error copying product details: $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    }
})

[void]$contextMenu.Items.Add($lookupPriceItem)
[void]$contextMenu.Items.Add($copyPricesItem)
[void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$contextMenu.Items.Add($deleteItem)
$itemsGrid.ContextMenuStrip = $contextMenu

# Handle cell value changes for auto-save
$itemsGrid.Add_CellEndEdit({
    param($control, $e)

    if ($script:AutoSaveEnabled -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
        $row = $itemsGrid.Rows[$e.RowIndex]
        $itemID = $row.Cells["ID"].Value
        $columnName = $itemsGrid.Columns[$e.ColumnIndex].Name
        $newValue = $row.Cells[$e.ColumnIndex].Value

        # Map display column names to database fields
        $fieldMap = @{
            "Lot" = "Lot"
            "Description" = "Description"
            "Brand" = "Brand"
            "Model" = "Model"
            "PartNumber" = "PartNumber"
            "Color" = "Color"
            "Category" = "Category"
            "Qty" = "Quantity"
            "PerItemCost" = "PerItemCost"
            "MSRP" = "CurrentMSRP"
            "CompetitorPrice" = "CurrentMarketAvg"
            "Our Price" = "SuggestedPrice"
        }

        if ($fieldMap.ContainsKey($columnName)) {
            $dbField = $fieldMap[$columnName]

            # Convert numeric fields to proper types
            $convertedValue = $newValue
            if ($columnName -in @("PerItemCost", "MSRP", "CompetitorPrice", "Our Price")) {
                try {
                    $convertedValue = [decimal]$newValue
                } catch {
                    $statusLabel.Text = "Error: Invalid number format for $columnName"
                    return
                }
            }
            elseif ($columnName -in @("Qty", "Lot")) {
                try {
                    $convertedValue = [int]$newValue
                } catch {
                    $statusLabel.Text = "Error: Invalid quantity - must be a whole number"
                    return
                }
            }

            # Save to database
            if (Update-ItemField -ItemID $itemID -FieldName $dbField -Value $convertedValue) {

                # Check if this is a field that requires recalculation
                if ($columnName -in @("PerItemCost", "MSRP", "CompetitorPrice", "Our Price", "Qty")) {
                    # Disable auto-save temporarily
                    $script:AutoSaveEnabled = $false

                    # Get ALL updated data for this item from database
                    $query = @"
                        SELECT
                            ItemID,
                            Lot,
                            Description,
                            Brand,
                            Model,
                            PartNumber,
                            Color,
                            Category,
                            ROUND(PerItemCost, 2) as PerItemCost,
                            ROUND(CurrentMSRP, 2) as MSRP,
                            ROUND(CurrentMarketAvg, 2) as CompetitorPrice,
                            ROUND(SuggestedPrice, 2) as OurPrice,
                            Quantity,
                            ROUND((SuggestedPrice - PerItemCost), 2) as PerItemProfit,
                            CASE WHEN PerItemCost > 0
                                THEN ROUND((SuggestedPrice/PerItemCost - 1) * 100, 1)
                                ELSE 0 END as Margin
                        FROM Items
                        WHERE ItemID = @ID
"@
                    $params = @{ ID = $itemID }
                    $result = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query -SqlParameters $params

                    if ($result) {
                        # Update ALL cells in the row to ensure consistency
                        # Handle null values from database
                        $row.Cells["Lot"].Value = if ($null -ne $result.Lot) { $result.Lot } else { 0 }
                        $row.Cells["Description"].Value = if ($null -ne $result.Description) { $result.Description } else { "" }
                        $row.Cells["Brand"].Value = if ($null -ne $result.Brand) { $result.Brand } else { "" }
                        $row.Cells["Model"].Value = if ($null -ne $result.Model) { $result.Model } else { "" }
                        $row.Cells["PartNumber"].Value = if ($null -ne $result.PartNumber) { $result.PartNumber } else { "" }
                        $row.Cells["Color"].Value = if ($null -ne $result.Color) { $result.Color } else { "" }
                        $row.Cells["Category"].Value = if ($null -ne $result.Category) { $result.Category } else { "" }
                        $row.Cells["PerItemCost"].Value = if ($null -ne $result.PerItemCost) { $result.PerItemCost } else { 0 }
                        $row.Cells["MSRP"].Value = if ($null -ne $result.MSRP) { $result.MSRP } else { 0 }
                        $row.Cells["CompetitorPrice"].Value = if ($null -ne $result.CompetitorPrice) { $result.CompetitorPrice } else { 0 }
                        $row.Cells["Our Price"].Value = if ($null -ne $result.OurPrice) { $result.OurPrice } else { 0 }
                        $row.Cells["Qty"].Value = if ($null -ne $result.Quantity) { $result.Quantity } else { 0 }
                        $row.Cells["PerItemProfit"].Value = if ($null -ne $result.PerItemProfit) { $result.PerItemProfit } else { 0 }
                        $row.Cells["Margin"].Value = if ($null -ne $result.Margin) { "$($result.Margin)%" } else { "0%" }

                        # Force the grid to redraw this row
                        $itemsGrid.InvalidateRow($e.RowIndex)
                        $itemsGrid.Update()
                        [System.Windows.Forms.Application]::DoEvents()

                        $statusLabel.Text = "UPDATED: Item #$itemID | Cost=$($result.PerItemCost) OurPrice=$($result.OurPrice) Profit=$($result.PerItemProfit) Margin=$($result.Margin)%"
                    }
                    else {
                        $statusLabel.Text = "WARNING: Could not retrieve recalculated values for Item #$itemID"
                    }

                    # Re-enable auto-save
                    $script:AutoSaveEnabled = $true
                }
                else {
                    # For non-numeric fields (Description, Brand, Model, Category, PartNumber, Color), just confirm save
                    $statusLabel.Text = "Saved: $columnName = '$convertedValue' for Item #$itemID"
                }

                # Update dashboard and charts
                Update-Dashboard
                Update-Charts
            }
            else {
                $statusLabel.Text = "ERROR: Failed to save $columnName for Item #$itemID"
            }
        }
    }
})

$itemsTab.Controls.AddRange(@($searchPanel, $itemsGrid))

# Analytics Tab
$analyticsTab = New-Object System.Windows.Forms.TabPage
$analyticsTab.Text = "Analytics"
$analyticsTab.BackColor = [System.Drawing.Color]::WhiteSmoke

# Create chart
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart.Location = New-Object System.Drawing.Point(10, 10)
$chart.Size = New-Object System.Drawing.Size(1140, 650)
$chart.BackColor = [System.Drawing.Color]::White

# Add chart area
$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartArea.Name = "MainArea"
$chartArea.BackColor = [System.Drawing.Color]::WhiteSmoke
$chart.ChartAreas.Add($chartArea)

# Add series for profit by category
$profitSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$profitSeries.Name = "Profit by Category"
$profitSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$profitSeries.Color = [System.Drawing.Color]::FromArgb(76, 175, 80)
$chart.Series.Add($profitSeries)

# Add title
$title = New-Object System.Windows.Forms.DataVisualization.Charting.Title
$title.Text = "Profit Analysis by Category"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$chart.Titles.Add($title)

# Add legend
$legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
$legend.Name = "Legend1"
$chart.Legends.Add($legend)

$analyticsTab.Controls.Add($chart)

# Add tabs to control
[void]$tabControl.Controls.Add($dashboardTab)
[void]$tabControl.Controls.Add($itemsTab)
[void]$tabControl.Controls.Add($analyticsTab)

# Status Bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready | Auto-Save: ON | Module: $(if($script:ModuleLoaded){'Loaded'}else{'Not Loaded'})"
[void]$statusBar.Items.Add($statusLabel)

# Add controls to form
$mainForm.Controls.AddRange(@($menuStrip, $tabControl, $statusBar))

#endregion

#region Helper Functions

function Update-Dashboard {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()

    try {
        $query = @"
            SELECT
                COUNT(*) as TotalItems,
                COALESCE(SUM(PerItemCost * Quantity), 0) as TotalInvestment,
                COALESCE(SUM(SuggestedPrice * Quantity), 0) as PotentialRevenue
            FROM Items WHERE IsActive = 1
"@
        $stats = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query

        if ($stats) {
            $lblTotalItemsValue.Text = $stats.TotalItems
            $lblInvestmentValue.Text = "$" + [Math]::Round($stats.TotalInvestment, 2)
            $lblRevenueValue.Text = "$" + [Math]::Round($stats.PotentialRevenue, 2)

            if ($stats.TotalInvestment -gt 0) {
                $margin = (($stats.PotentialRevenue - $stats.TotalInvestment) / $stats.TotalInvestment) * 100
                $lblMarginValue.Text = [Math]::Round($margin, 1).ToString() + "%"
            }
        }
    }
    catch {
        $statusLabel.Text = "Error updating dashboard: $_"
    }
}

function Update-ItemGrid {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param([string]$Filter = "")

    try {
        $script:AutoSaveEnabled = $false

        $items = Get-PurchaseItemLocal -Filter $Filter

        # Create DataTable
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("ID", [int])
        [void]$dt.Columns.Add("Lot", [int])
        [void]$dt.Columns.Add("Description", [string])
        [void]$dt.Columns.Add("Brand", [string])
        [void]$dt.Columns.Add("Model", [string])
        [void]$dt.Columns.Add("PartNumber", [string])
        [void]$dt.Columns.Add("Color", [string])
        [void]$dt.Columns.Add("Category", [string])
        [void]$dt.Columns.Add("Qty", [int])
        [void]$dt.Columns.Add("PerItemCost", [decimal])
        [void]$dt.Columns.Add("MSRP", [decimal])
        [void]$dt.Columns.Add("CompetitorPrice", [decimal])
        [void]$dt.Columns.Add("Our Price", [decimal])
        [void]$dt.Columns.Add("PerItemProfit", [decimal])
        [void]$dt.Columns.Add("Margin", [string])

        foreach ($item in $items) {
            $row = $dt.NewRow()
            $row["ID"] = $item.ItemID
            $row["Lot"] = if ($item.Lot) { $item.Lot } else { 0 }
            $row["Description"] = if ($item.Description) { $item.Description } else { "" }
            $row["Brand"] = if ($item.Brand) { $item.Brand } else { "" }
            $row["Model"] = if ($item.Model) { $item.Model } else { "" }
            $row["PartNumber"] = if ($item.PartNumber) { $item.PartNumber } else { "" }
            $row["Color"] = if ($item.Color) { $item.Color } else { "" }
            $row["Category"] = if ($item.Category) { $item.Category } else { "" }
            $row["Qty"] = if ($item.Quantity) { $item.Quantity } else { 0 }
            $row["PerItemCost"] = if ($item.PerItemCost) { $item.PerItemCost } else { 0 }
            $row["MSRP"] = if ($item.MSRP) { $item.MSRP } else { 0 }
            $row["CompetitorPrice"] = if ($item.CompetitorPrice) { $item.CompetitorPrice } else { 0 }
            $row["Our Price"] = if ($item.OurPrice) { $item.OurPrice } else { 0 }
            $row["PerItemProfit"] = if ($item.PerItemProfit) { $item.PerItemProfit } else { 0 }
            $row["Margin"] = if ($item.Margin) { "$($item.Margin)%" } else { "0%" }
            [void]$dt.Rows.Add($row)
        }

        $itemsGrid.DataSource = $dt

        # Format columns
        if ($itemsGrid.Columns.Count -gt 0) {
            $itemsGrid.Columns["ID"].ReadOnly = $true
            $itemsGrid.Columns["ID"].Width = 50
            $itemsGrid.Columns["Lot"].ReadOnly = $true
            $itemsGrid.Columns["Lot"].Width = 60
            $itemsGrid.Columns["Description"].Width = 180
            $itemsGrid.Columns["Brand"].Width = 90
            $itemsGrid.Columns["Model"].Width = 90
            $itemsGrid.Columns["PartNumber"].Width = 100
            $itemsGrid.Columns["Color"].Width = 70
            $itemsGrid.Columns["Category"].Width = 80
            $itemsGrid.Columns["Qty"].Width = 50
            $itemsGrid.Columns["PerItemCost"].Width = 90
            $itemsGrid.Columns["MSRP"].Width = 70
            $itemsGrid.Columns["CompetitorPrice"].Width = 110
            $itemsGrid.Columns["Our Price"].Width = 80
            $itemsGrid.Columns["PerItemProfit"].ReadOnly = $true
            $itemsGrid.Columns["PerItemProfit"].Width = 90
            $itemsGrid.Columns["Margin"].ReadOnly = $true
            $itemsGrid.Columns["Margin"].Width = 70
        }

        $statusLabel.Text = "Found $($items.Count) items | Auto-Save: ON"
        $script:AutoSaveEnabled = $true
    }
    catch {
        $statusLabel.Text = "Error loading items: $_"
        $script:AutoSaveEnabled = $true
    }
}

function Update-Charts {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
    param()

    try {
        $query = @"
            SELECT
                Category,
                SUM((SuggestedPrice - PerItemCost) * Quantity) as TotalProfit,
                COUNT(*) as ItemCount
            FROM Items
            WHERE IsActive = 1
            GROUP BY Category
            ORDER BY TotalProfit DESC
"@
        $data = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query

        $chart.Series["Profit by Category"].Points.Clear()

        foreach ($row in $data) {
            $category = if ($row.Category) { $row.Category } else { "Uncategorized" }
            $profit = if ($row.TotalProfit) { [Math]::Round($row.TotalProfit, 2) } else { 0 }

            $point = $chart.Series["Profit by Category"].Points.AddXY($category, $profit)
            $chart.Series["Profit by Category"].Points[$point].Label = "`$$profit"
        }
    }
    catch {
        Write-Verbose "Chart update failed: $_"
    }
}

function Show-AddItemDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Add New Item"
    $dialog.Size = New-Object System.Drawing.Size(400, 300)
    $dialog.StartPosition = "CenterParent"

    # Description
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Description:"
    $lblDesc.Location = New-Object System.Drawing.Point(10, 20)
    $lblDesc.Size = New-Object System.Drawing.Size(100, 20)

    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Location = New-Object System.Drawing.Point(120, 20)
    $txtDesc.Size = New-Object System.Drawing.Size(250, 20)

    # Category
    $lblCat = New-Object System.Windows.Forms.Label
    $lblCat.Text = "Category:"
    $lblCat.Location = New-Object System.Drawing.Point(10, 50)
    $lblCat.Size = New-Object System.Drawing.Size(100, 20)

    $txtCat = New-Object System.Windows.Forms.TextBox
    $txtCat.Location = New-Object System.Drawing.Point(120, 50)
    $txtCat.Size = New-Object System.Drawing.Size(250, 20)
    $txtCat.Text = "General"

    # Quantity
    $lblQty = New-Object System.Windows.Forms.Label
    $lblQty.Text = "Quantity:"
    $lblQty.Location = New-Object System.Drawing.Point(10, 80)
    $lblQty.Size = New-Object System.Drawing.Size(100, 20)

    $txtQty = New-Object System.Windows.Forms.TextBox
    $txtQty.Location = New-Object System.Drawing.Point(120, 80)
    $txtQty.Size = New-Object System.Drawing.Size(100, 20)
    $txtQty.Text = "1"

    # Cost
    $lblCost = New-Object System.Windows.Forms.Label
    $lblCost.Text = "Per Item Cost:"
    $lblCost.Location = New-Object System.Drawing.Point(10, 110)
    $lblCost.Size = New-Object System.Drawing.Size(100, 20)

    $txtCost = New-Object System.Windows.Forms.TextBox
    $txtCost.Location = New-Object System.Drawing.Point(120, 110)
    $txtCost.Size = New-Object System.Drawing.Size(100, 20)

    # Our Price
    $lblPrice = New-Object System.Windows.Forms.Label
    $lblPrice.Text = "Our Price:"
    $lblPrice.Location = New-Object System.Drawing.Point(10, 140)
    $lblPrice.Size = New-Object System.Drawing.Size(100, 20)

    $txtPrice = New-Object System.Windows.Forms.TextBox
    $txtPrice.Location = New-Object System.Drawing.Point(120, 140)
    $txtPrice.Size = New-Object System.Drawing.Size(100, 20)

    # Buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Add"
    $btnOK.Location = New-Object System.Drawing.Point(120, 200)
    $btnOK.Size = New-Object System.Drawing.Size(75, 23)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(200, 200)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 23)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dialog.Controls.AddRange(@(
        $lblDesc, $txtDesc, $lblCat, $txtCat,
        $lblQty, $txtQty, $lblCost, $txtCost,
        $lblPrice, $txtPrice, $btnOK, $btnCancel
    ))

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $query = @"
                INSERT INTO Items (Description, Category, Quantity, PerItemCost, SuggestedPrice)
                VALUES (@Desc, @Cat, @Qty, @Cost, @Price)
"@
            $params = @{
                Desc = $txtDesc.Text
                Cat = $txtCat.Text
                Qty = [int]$txtQty.Text
                Cost = if ($txtCost.Text) { [decimal]$txtCost.Text } else { 0 }
                Price = if ($txtPrice.Text) { [decimal]$txtPrice.Text } else { 0 }
            }

            Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query -SqlParameters $params

            Update-Dashboard
            Update-ItemGrid
            Update-Charts

            [System.Windows.Forms.MessageBox]::Show("Item added successfully!", "Success")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error")
        }
    }
}

function Import-CSVData {
    param([string]$Path)

    try {
        $data = Import-Csv -Path $Path
        $count = 0

        foreach ($row in $data) {
            $query = @"
                INSERT INTO Items (Description, Category, Quantity, PerItemCost, SuggestedPrice)
                VALUES (@Desc, @Cat, @Qty, @Cost, @Price)
"@
            $params = @{
                Desc = $row.Description
                Cat = if ($row.Category) { $row.Category } else { "General" }
                Qty = if ($row.Quantity) { [int]$row.Quantity } else { 1 }
                Cost = if ($row.'Per Item Cost') {
                    [decimal]($row.'Per Item Cost' -replace '[$,]', '')
                } else { 0 }
                Price = if ($row.'Our Asking Price') {
                    [decimal]($row.'Our Asking Price' -replace '[$,]', '')
                } else { 0 }
            }

            Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query -SqlParameters $params
            $count++
        }

        Update-Dashboard
        Update-ItemGrid
        Update-Charts

        [System.Windows.Forms.MessageBox]::Show(
            "Imported $count items successfully!",
            "Import Complete"
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Import Error: $_", "Error")
    }
}

function Export-CSVData {
    param([string]$Path)

    try {
        $query = "SELECT * FROM Items WHERE IsActive = 1 ORDER BY ItemID"
        $data = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $query
        $data | Export-Csv -Path $Path -NoTypeInformation

        [System.Windows.Forms.MessageBox]::Show(
            "Data exported successfully!",
            "Export Complete"
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Export Error: $_", "Error")
    }
}

#endregion

# Initialize and show form
Initialize-Database
Update-Dashboard
Update-ItemGrid
Update-Charts

# Show the form
[void]$mainForm.ShowDialog()