<#
.SYNOPSIS
    Flexible database schema updater for adding, renaming, deleting columns, and changing column types.

.DESCRIPTION
    This script provides flexible database schema management:
    - Add new columns with specified data types
    - Rename existing columns (by creating new and migrating data)
    - Delete columns (with safety confirmation)
    - Change column data types (with safety confirmation)
    - Automatically calculate values for new columns
    - Create helpful views for analysis

.PARAMETER DatabasePath
    Path to the SQLite database file to update.

.PARAMETER TableName
    Name of the table to modify. Default is "Items".

.PARAMETER AddColumns
    Hashtable of columns to add. Format: @{ColumnName="DataType"; ColumnName2="DataType2"}
    Example: @{MarkupPercentage="REAL"; TotalProfit="REAL DEFAULT 0"}

.PARAMETER RenameColumns
    Hashtable mapping old column names to new names. Format: @{OldName="NewName"; OldName2="NewName2"}
    Example: @{CurrentMarketAvg="CompetitorPrice"}

.PARAMETER DeleteColumns
    Array of column names to delete. WARNING: This is destructive and requires table recreation.
    Example: @("ObsoleteColumn1", "OldColumn2")

.PARAMETER ChangeColumnTypes
    Hashtable of columns to change types. Format: @{ColumnName="NewType"}
    Example: @{PartNumber="TEXT"; Quantity="INTEGER"}
    WARNING: This is destructive and requires table recreation.

.PARAMETER CalculateValues
    Switch to automatically calculate values for specific known columns (MarkupPercentage, TotalProfit).

.PARAMETER CreateViews
    Switch to create helpful analysis views after schema changes.

.PARAMETER BackupFirst
    Creates a backup before making changes (strongly recommended).

.EXAMPLE
    # Add new columns
    .\Update-DatabaseSchema.ps1 -DatabasePath ".\purchasesdata.db" `
        -AddColumns @{MarkupPercentage="REAL"; TotalProfit="REAL DEFAULT 0"} `
        -BackupFirst

.EXAMPLE
    # Rename a column
    .\Update-DatabaseSchema.ps1 -DatabasePath ".\purchasesdata.db" `
        -RenameColumns @{CurrentMarketAvg="CompetitorPrice"} `
        -BackupFirst

.EXAMPLE
    # Change PartNumber from numeric to TEXT
    .\Update-DatabaseSchema.ps1 -DatabasePath ".\purchasesdata.db" `
        -ChangeColumnTypes @{PartNumber="TEXT"} `
        -BackupFirst

.EXAMPLE
    # Add and rename in one operation
    .\Update-DatabaseSchema.ps1 -DatabasePath ".\purchasesdata.db" `
        -AddColumns @{ProfitMargin="REAL"} `
        -RenameColumns @{OldPrice="HistoricalPrice"} `
        -CalculateValues -CreateViews -BackupFirst

.EXAMPLE
    # Delete columns (requires confirmation)
    .\Update-DatabaseSchema.ps1 -DatabasePath ".\purchasesdata.db" `
        -DeleteColumns @("ObsoleteColumn", "UnusedField") `
        -BackupFirst

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10/10/2025
    Version: 3.1.2
    Change Date: 10/14/2025
    Change Purpose: Fixed transaction commit timing

.CHANGELOG
    3.1.2 - 10/14/2025
        - Fixed transaction commit timing to occur before view recreation
        - Eliminated "cannot commit - no transaction is active" error message
    3.1.1 - 10/14/2025
        - Fixed Update-ColumnType to drop and recreate dependent views
        - Added better error handling for transaction rollback
        - Now detects and handles views that reference the table being modified
    3.1.0 - 10/14/2025
        - Added Update-ColumnType function to change column data types
        - Added -ChangeColumnTypes parameter
        - Uses table recreation with transaction safety
        - Automatic type conversion during data migration
    3.0.1 - 10/13/2025
        - Fixed PSScriptAnalyzer ShouldProcess warnings
        - Added [CmdletBinding(SupportsShouldProcess)] to all state-changing functions
    3.0.0 - 10/13/2025
        - Added parameterization for flexible column operations
        - Added -AddColumns parameter to add custom columns
        - Added -RenameColumns parameter to rename existing columns
        - Added -DeleteColumns parameter to delete columns (with warnings)
        - Added -TableName parameter to specify target table
        - Added -CalculateValues switch for automatic value calculation
        - Added -CreateViews switch for view creation
        - Made script more modular and reusable
    2.0.0 - 10/13/2025
        - Fixed all PSScriptAnalyzer warnings
        - Replaced Write-Host with Write-Information
        - Renamed functions to use approved verbs and singular nouns
        - Removed unused variables
    1.0.0 - 10/10/2025
        - Initial release
        - Add new columns and rename existing column
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
    [string]$DatabasePath,

    [Parameter(Mandatory=$false)]
    [string]$TableName = "Items",

    [Parameter(Mandatory=$false)]
    [hashtable]$AddColumns,

    [Parameter(Mandatory=$false)]
    [hashtable]$RenameColumns,

    [Parameter(Mandatory=$false)]
    [string[]]$DeleteColumns,

    [Parameter(Mandatory=$false)]
    [hashtable]$ChangeColumnTypes,

    [Parameter(Mandatory=$false)]
    [switch]$CalculateValues,

    [Parameter(Mandatory=$false)]
    [switch]$CreateViews,

    [Parameter(Mandatory=$false)]
    [switch]$BackupFirst
)

# Import SQLite module
try {
    Import-Module PSSQLite -ErrorAction Stop
}
catch {
    Write-Error "PSSQLite module not found. Please install: Install-Module PSSQLite -Scope CurrentUser"
    exit 1
}

Write-Information "`n===============================================" -InformationAction Continue
Write-Information "     DATABASE SCHEMA UPDATE TOOL" -InformationAction Continue
Write-Information "===============================================" -InformationAction Continue
Write-Information "Database: $DatabasePath" -InformationAction Continue
Write-Information "Table: $TableName" -InformationAction Continue

#region Backup Function

function Backup-Database {
    param([string]$DbPath)

    $backupPath = $DbPath -replace '\.db$', "_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').db"

    Write-Information "`n📁 Creating backup..." -InformationAction Continue
    try {
        Copy-Item -Path $DbPath -Destination $backupPath -Force
        Write-Information "✅ Backup created: $backupPath" -InformationAction Continue
        return $backupPath
    }
    catch {
        Write-Error "Failed to create backup: $_"
        $response = Read-Host "Continue without backup? (yes/no)"
        if ($response -ne 'yes') {
            exit 1
        }
    }
}

#endregion

#region Schema Check Functions

function Get-CurrentSchema {
    param(
        [string]$DbPath,
        [string]$Table
    )

    Write-Information "`n🔍 Checking current schema..." -InformationAction Continue

    $query = "PRAGMA table_info($Table)"
    $columns = Invoke-SqliteQuery -DataSource $DbPath -Query $query

    Write-Information "`nCurrent columns in $Table table:" -InformationAction Continue
    foreach ($col in $columns) {
        Write-Information "  - $($col.name) ($($col.type))" -InformationAction Continue
    }

    return $columns
}

function Test-Column {
    param(
        [string]$DbPath,
        [string]$Table,
        [string]$ColumnName
    )

    $query = "PRAGMA table_info($Table)"
    $columns = Invoke-SqliteQuery -DataSource $DbPath -Query $query

    return ($columns | Where-Object { $_.name -eq $ColumnName }).Count -gt 0
}

#endregion

#region Add Column Functions

function Add-NewColumn {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$DbPath,
        [string]$Table,
        [hashtable]$Columns
    )

    if (-not $Columns -or $Columns.Count -eq 0) {
        Write-Information "⚠️  No columns to add" -InformationAction Continue
        return
    }

    Write-Information "`n📝 Adding new columns..." -InformationAction Continue

    foreach ($columnName in $Columns.Keys) {
        $dataType = $Columns[$columnName]

        if (Test-Column -DbPath $DbPath -Table $Table -ColumnName $columnName) {
            Write-Information "⚠️  Column '$columnName' already exists, skipping" -InformationAction Continue
            continue
        }

        try {
            $query = "ALTER TABLE $Table ADD COLUMN $columnName $dataType"
            if ($PSCmdlet.ShouldProcess("$Table table", "Add column $columnName ($dataType)")) {
                Invoke-SqliteQuery -DataSource $DbPath -Query $query
                Write-Information "✅ Added column: $columnName ($dataType)" -InformationAction Continue
            }
        }
        catch {
            Write-Warning "Could not add $columnName column: $_"
        }
    }
}

#endregion

#region Rename Column Functions

function Move-ColumnData {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$DbPath,
        [string]$Table,
        [hashtable]$RenameMap
    )

    if (-not $RenameMap -or $RenameMap.Count -eq 0) {
        Write-Information "⚠️  No columns to rename" -InformationAction Continue
        return
    }

    Write-Information "`n🔄 Renaming columns (SQLite workaround: create new + migrate data)..." -InformationAction Continue

    foreach ($oldName in $RenameMap.Keys) {
        $newName = $RenameMap[$oldName]

        # Check if old column exists
        if (-not (Test-Column -DbPath $DbPath -Table $Table -ColumnName $oldName)) {
            Write-Warning "Column '$oldName' does not exist, skipping rename"
            continue
        }

        # Get the data type of the old column
        $query = "PRAGMA table_info($Table)"
        $columns = Invoke-SqliteQuery -DataSource $DbPath -Query $query
        $oldColumn = $columns | Where-Object { $_.name -eq $oldName }
        $dataType = $oldColumn.type

        try {
            # Add new column if it doesn't exist
            if (-not (Test-Column -DbPath $DbPath -Table $Table -ColumnName $newName)) {
                $addQuery = "ALTER TABLE $Table ADD COLUMN $newName $dataType"
                if ($PSCmdlet.ShouldProcess("$Table table", "Add column $newName ($dataType)")) {
                    Invoke-SqliteQuery -DataSource $DbPath -Query $addQuery
                    Write-Information "✅ Created new column: $newName" -InformationAction Continue
                }
            }

            # Copy data from old to new
            $copyQuery = "UPDATE $Table SET $newName = $oldName WHERE $oldName IS NOT NULL AND ($newName IS NULL OR $newName = 0)"
            if ($PSCmdlet.ShouldProcess("$Table table", "Copy data from $oldName to $newName")) {
                Invoke-SqliteQuery -DataSource $DbPath -Query $copyQuery
                Write-Information "✅ Data migrated from $oldName → $newName" -InformationAction Continue
            }

            Write-Information "ℹ️  Note: Old column '$oldName' still exists. Use -DeleteColumns to remove it if desired." -InformationAction Continue
        }
        catch {
            Write-Warning "Could not rename $oldName to $newName : $_"
        }
    }
}

#endregion

#region Delete Column Functions

function Remove-TableColumn {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$DbPath,
        [string]$Table,
        [string[]]$Columns
    )

    if (-not $Columns -or $Columns.Count -eq 0) {
        Write-Information "⚠️  No columns to delete" -InformationAction Continue
        return
    }

    Write-Warning "`n⚠️  DESTRUCTIVE OPERATION: Deleting columns requires table recreation!"
    Write-Information "Columns to delete: $($Columns -join ', ')" -InformationAction Continue

    $confirmation = Read-Host "`nType 'DELETE' to confirm deletion of these columns"
    if ($confirmation -ne 'DELETE') {
        Write-Information "❌ Column deletion cancelled" -InformationAction Continue
        return
    }

    Write-Information "`n🗑️  Deleting columns (SQLite workaround: recreate table)..." -InformationAction Continue

    try {
        # Get current table schema
        $pragmaQuery = "PRAGMA table_info($Table)"
        $currentColumns = Invoke-SqliteQuery -DataSource $DbPath -Query $pragmaQuery

        # Filter out columns to delete
        $keepColumns = $currentColumns | Where-Object { $Columns -notcontains $_.name }

        if ($keepColumns.Count -eq 0) {
            Write-Error "Cannot delete all columns from table!"
            return
        }

        # Build column definitions for new table
        $columnDefs = $keepColumns | ForEach-Object {
            $def = "$($_.name) $($_.type)"
            if ($_.notnull -eq 1) { $def += " NOT NULL" }
            if ($_.dflt_value) { $def += " DEFAULT $($_.dflt_value)" }
            if ($_.pk -eq 1) { $def += " PRIMARY KEY" }
            $def
        }

        $columnDefString = $columnDefs -join ", "
        $keepColumnNames = ($keepColumns | ForEach-Object { $_.name }) -join ", "

        if ($PSCmdlet.ShouldProcess("$Table table", "Recreate table without columns: $($Columns -join ', ')")) {
            # Create new table
            $createQuery = "CREATE TABLE ${Table}_new ($columnDefString)"
            Invoke-SqliteQuery -DataSource $DbPath -Query $createQuery

            # Copy data
            $copyQuery = "INSERT INTO ${Table}_new ($keepColumnNames) SELECT $keepColumnNames FROM $Table"
            Invoke-SqliteQuery -DataSource $DbPath -Query $copyQuery

            # Drop old table
            $dropQuery = "DROP TABLE $Table"
            Invoke-SqliteQuery -DataSource $DbPath -Query $dropQuery

            # Rename new table
            $renameQuery = "ALTER TABLE ${Table}_new RENAME TO $Table"
            Invoke-SqliteQuery -DataSource $DbPath -Query $renameQuery

            Write-Information "✅ Columns deleted successfully: $($Columns -join ', ')" -InformationAction Continue
        }
    }
    catch {
        Write-Error "Failed to delete columns: $_"
        Write-Error "You may need to restore from backup!"
    }
}

#endregion

#region Change Column Type Functions

function Update-ColumnType {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DbPath,

        [Parameter(Mandatory=$true)]
        [string]$Table,

        [Parameter(Mandatory=$true)]
        [hashtable]$ColumnTypes
    )

    if (-not $ColumnTypes -or $ColumnTypes.Count -eq 0) {
        Write-Information "⚠️  No column types to change" -InformationAction Continue
        return
    }

    Write-Warning "`n⚠️  DESTRUCTIVE OPERATION: Changing column types requires table recreation!"
    Write-Information "Columns to modify:" -InformationAction Continue
    foreach ($col in $ColumnTypes.Keys) {
        Write-Information "  - $col → $($ColumnTypes[$col])" -InformationAction Continue
    }

    $confirmation = Read-Host "`nType 'CHANGE' to confirm type changes for these columns"
    if ($confirmation -ne 'CHANGE') {
        Write-Information "❌ Column type change cancelled" -InformationAction Continue
        return
    }

    Write-Information "`n🔄 Changing column types (SQLite workaround: recreate table)..." -InformationAction Continue

    try {
        # Get current table schema
        $pragmaQuery = "PRAGMA table_info($Table)"
        $currentColumns = Invoke-SqliteQuery -DataSource $DbPath -Query $pragmaQuery

        if ($currentColumns.Count -eq 0) {
            Write-Error "Table '$Table' does not exist or has no columns!"
            return
        }

        # Get list of dependent views that need to be dropped and recreated
        $viewQuery = @"
SELECT name, sql FROM sqlite_master
WHERE type = 'view'
AND sql LIKE '%$Table%'
"@
        $dependentViews = Invoke-SqliteQuery -DataSource $DbPath -Query $viewQuery

        if ($dependentViews -and $dependentViews.Count -gt 0) {
            Write-Information "  📋 Found $($dependentViews.Count) dependent view(s) that will be recreated" -InformationAction Continue
        }

        # Build column definitions for new table with updated types
        $columnDefs = $currentColumns | ForEach-Object {
            $colName = $_.name

            # Use new type if specified, otherwise keep original
            $colType = if ($ColumnTypes.ContainsKey($colName)) {
                $ColumnTypes[$colName]
                Write-Information "  ✏️  Changing $colName from $($_.type) to $($ColumnTypes[$colName])" -InformationAction Continue
            } else {
                $_.type
            }

            $def = "$colName $colType"
            if ($_.notnull -eq 1) { $def += " NOT NULL" }
            if ($_.dflt_value) { $def += " DEFAULT $($_.dflt_value)" }
            if ($_.pk -eq 1) { $def += " PRIMARY KEY" }
            $def
        }

        $columnDefString = $columnDefs -join ", "
        $columnNames = ($currentColumns | ForEach-Object { $_.name }) -join ", "

        if ($PSCmdlet.ShouldProcess("$Table table", "Change column types: $($ColumnTypes.Keys -join ', ')")) {
            # Begin transaction for safety
            Invoke-SqliteQuery -DataSource $DbPath -Query "BEGIN TRANSACTION"

            try {
                # Drop dependent views first
                if ($dependentViews) {
                    foreach ($view in $dependentViews) {
                        Write-Information "  🗑️  Dropping view: $($view.name)" -InformationAction Continue
                        Invoke-SqliteQuery -DataSource $DbPath -Query "DROP VIEW IF EXISTS $($view.name)"
                    }
                }

                # Create new table with updated schema
                $createQuery = "CREATE TABLE ${Table}_new ($columnDefString)"
                Invoke-SqliteQuery -DataSource $DbPath -Query $createQuery

                # Copy data with type conversion
                # SQLite will automatically convert types where possible
                $copyQuery = "INSERT INTO ${Table}_new ($columnNames) SELECT $columnNames FROM $Table"
                Invoke-SqliteQuery -DataSource $DbPath -Query $copyQuery

                # Drop old table
                $dropQuery = "DROP TABLE $Table"
                Invoke-SqliteQuery -DataSource $DbPath -Query $dropQuery

                # Rename new table
                $renameQuery = "ALTER TABLE ${Table}_new RENAME TO $Table"
                Invoke-SqliteQuery -DataSource $DbPath -Query $renameQuery

                # Commit transaction before recreating views
                # (views cannot be created inside a transaction in some SQLite configurations)
                Invoke-SqliteQuery -DataSource $DbPath -Query "COMMIT"

                # Recreate dependent views
                if ($dependentViews) {
                    foreach ($view in $dependentViews) {
                        Write-Information "  ✅ Recreating view: $($view.name)" -InformationAction Continue
                        Invoke-SqliteQuery -DataSource $DbPath -Query $view.sql
                    }
                }

                Write-Information "✅ Column types changed successfully!" -InformationAction Continue

                # Show updated schema
                $updatedColumns = Invoke-SqliteQuery -DataSource $DbPath -Query $pragmaQuery
                Write-Information "`nUpdated column types:" -InformationAction Continue
                foreach ($col in $updatedColumns | Where-Object { $ColumnTypes.ContainsKey($_.name) }) {
                    Write-Information "  ✅ $($col.name): $($col.type)" -InformationAction Continue
                }
            }
            catch {
                # Rollback on error
                try {
                    Invoke-SqliteQuery -DataSource $DbPath -Query "ROLLBACK"
                    Write-Warning "Transaction rolled back due to error"
                }
                catch {
                    Write-Warning "Could not rollback transaction: $_"
                }
                throw
            }
        }
    }
    catch {
        Write-Error "Failed to change column types: $_"
        Write-Error "You may need to restore from backup!"
    }
}

#endregion

#region Value Calculation Functions

function Update-CalculatedValue {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$DbPath,
        [string]$Table
    )

    Write-Information "`n📊 Calculating values for known columns..." -InformationAction Continue

    # Calculate MarkupPercentage if column exists
    if (Test-Column -DbPath $DbPath -Table $Table -ColumnName "MarkupPercentage") {
        try {
            $query = @"
                UPDATE $Table
                SET MarkupPercentage = ROUND(
                    CASE
                        WHEN PerItemCost > 0 AND SuggestedPrice > 0
                        THEN ((SuggestedPrice - PerItemCost) / PerItemCost) * 100
                        ELSE 0
                    END, 2
                )
                WHERE IsActive = 1
"@
            if ($PSCmdlet.ShouldProcess("$Table table", "Calculate MarkupPercentage values")) {
                Invoke-SqliteQuery -DataSource $DbPath -Query $query
                Write-Information "✅ Calculated MarkupPercentage values" -InformationAction Continue
            }
        }
        catch {
            Write-Warning "Error calculating MarkupPercentage: $_"
        }
    }

    # Calculate TotalProfit if column exists
    if (Test-Column -DbPath $DbPath -Table $Table -ColumnName "TotalProfit") {
        try {
            $query = @"
                UPDATE $Table
                SET TotalProfit = ROUND(
                    CASE
                        WHEN SuggestedPrice > 0 AND PerItemCost > 0 AND Quantity > 0
                        THEN (SuggestedPrice - PerItemCost) * Quantity
                        ELSE 0
                    END, 2
                )
                WHERE IsActive = 1
"@
            if ($PSCmdlet.ShouldProcess("$Table table", "Calculate TotalProfit values")) {
                Invoke-SqliteQuery -DataSource $DbPath -Query $query
                Write-Information "✅ Calculated TotalProfit values" -InformationAction Continue
            }
        }
        catch {
            Write-Warning "Error calculating TotalProfit: $_"
        }
    }
}

#endregion

#region View Creation Functions

function New-AnalysisView {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$DbPath,
        [string]$Table
    )

    Write-Information "`n📐 Creating analysis view..." -InformationAction Continue

    # Check if required columns exist for the view
    $hasCompetitorPrice = Test-Column -DbPath $DbPath -Table $Table -ColumnName "CompetitorPrice"
    $hasMarkupPercentage = Test-Column -DbPath $DbPath -Table $Table -ColumnName "MarkupPercentage"
    $hasTotalProfit = Test-Column -DbPath $DbPath -Table $Table -ColumnName "TotalProfit"

    if (-not ($hasCompetitorPrice -or $hasMarkupPercentage -or $hasTotalProfit)) {
        Write-Information "⚠️  No profit analysis columns found, skipping view creation" -InformationAction Continue
        return
    }

    try {
        # Drop existing view if it exists
        $dropQuery = "DROP VIEW IF EXISTS ItemProfitAnalysis"
        Invoke-SqliteQuery -DataSource $DbPath -Query $dropQuery

        # Build view query based on available columns
        $selectClause = @"
            ItemID,
            Description,
            Category,
            Quantity,
            PerItemCost,
            SuggestedPrice as OurPrice
"@

        if ($hasCompetitorPrice) {
            $selectClause += ",`n            CompetitorPrice"
        }
        if ($hasMarkupPercentage) {
            $selectClause += ",`n            MarkupPercentage"
        }
        if ($hasTotalProfit) {
            $selectClause += ",`n            TotalProfit"
        }

        $selectClause += ",`n            ROUND((SuggestedPrice - PerItemCost), 2) as ProfitPerItem"

        if ($hasCompetitorPrice) {
            $selectClause += @"
,
            CASE
                WHEN CompetitorPrice > 0
                THEN ROUND(((CompetitorPrice - SuggestedPrice) / CompetitorPrice) * 100, 2)
                ELSE 0
            END as PriceDifferencePercent,
            CASE
                WHEN SuggestedPrice < CompetitorPrice THEN 'Below Market'
                WHEN SuggestedPrice > CompetitorPrice THEN 'Above Market'
                ELSE 'At Market'
            END as PricePosition
"@
        }

        $viewQuery = @"
            CREATE VIEW ItemProfitAnalysis AS
            SELECT
$selectClause
            FROM $Table
            WHERE IsActive = 1
"@

        if ($PSCmdlet.ShouldProcess("Database", "Create ItemProfitAnalysis view")) {
            Invoke-SqliteQuery -DataSource $DbPath -Query $viewQuery
            Write-Information "✅ Created view: ItemProfitAnalysis" -InformationAction Continue
        }
    }
    catch {
        Write-Warning "Could not create view: $_"
    }
}

#endregion

#region Display Functions

function Show-UpdatedData {
    param(
        [string]$DbPath,
        [string]$Table
    )

    Write-Information "`n📋 Sample of updated data:" -InformationAction Continue

    try {
        # Build dynamic query based on available columns
        $columns = Invoke-SqliteQuery -DataSource $DbPath -Query "PRAGMA table_info($Table)"

        if (-not $columns -or $columns.Count -eq 0) {
            Write-Warning "No columns found in table $Table"
            return
        }

        $columnNames = ($columns | Select-Object -First 8 | ForEach-Object { $_.name }) -join ", "

        if ([string]::IsNullOrWhiteSpace($columnNames)) {
            Write-Warning "Could not build column list for query"
            return
        }

        $query = "SELECT $columnNames FROM $Table WHERE IsActive = 1 LIMIT 5"

        $data = Invoke-SqliteQuery -DataSource $DbPath -Query $query
        if ($data) {
            $data | Format-Table -AutoSize
        }
        else {
            Write-Information "No data to display" -InformationAction Continue
        }
    }
    catch {
        Write-Warning "Could not retrieve sample data: $_"
    }
}

#endregion

#region Main Execution

try {
    # Validate that at least one operation is requested
    if (-not $AddColumns -and -not $RenameColumns -and -not $DeleteColumns -and -not $ChangeColumnTypes -and -not $CalculateValues -and -not $CreateViews) {
        Write-Error "No operations specified. Use -AddColumns, -RenameColumns, -DeleteColumns, -ChangeColumnTypes, -CalculateValues, or -CreateViews"
        exit 1
    }

    # Create backup if requested
    $backupPath = $null
    if ($BackupFirst) {
        $backupPath = Backup-Database -DbPath $DatabasePath
    }

    # Show current schema
    $null = Get-CurrentSchema -DbPath $DatabasePath -Table $TableName

    # Perform requested operations
    if ($AddColumns) {
        Add-NewColumn -DbPath $DatabasePath -Table $TableName -Columns $AddColumns
    }

    if ($RenameColumns) {
        Move-ColumnData -DbPath $DatabasePath -Table $TableName -RenameMap $RenameColumns
    }

    if ($ChangeColumnTypes) {
        Update-ColumnType -DbPath $DatabasePath -Table $TableName -ColumnTypes $ChangeColumnTypes
    }

    if ($DeleteColumns) {
        Remove-TableColumn -DbPath $DatabasePath -Table $TableName -Columns $DeleteColumns
    }

    if ($CalculateValues) {
        Update-CalculatedValue -DbPath $DatabasePath -Table $TableName
    }

    if ($CreateViews) {
        New-AnalysisView -DbPath $DatabasePath -Table $TableName
    }

    # Show updated schema
    Write-Information "`n✨ Updated Schema:" -InformationAction Continue
    $null = Get-CurrentSchema -DbPath $DatabasePath -Table $TableName

    # Show sample data
    Show-UpdatedData -DbPath $DatabasePath -Table $TableName

    # Summary
    Write-Information "`n===============================================" -InformationAction Continue
    Write-Information "     UPDATE COMPLETED SUCCESSFULLY!" -InformationAction Continue
    Write-Information "===============================================" -InformationAction Continue

    if ($backupPath) {
        Write-Information "`n💾 Backup saved at: $backupPath" -InformationAction Continue
    }

}
catch {
    Write-Information "`n❌ ERROR: $_" -InformationAction Continue
    Write-Information "Stack trace: $($_.ScriptStackTrace)" -InformationAction Continue

    if ($backupPath) {
        Write-Information "`n🔄 To restore from backup, run:" -InformationAction Continue
        Write-Information "  Copy-Item '$backupPath' '$DatabasePath' -Force" -InformationAction Continue
    }

    throw
}

#endregion