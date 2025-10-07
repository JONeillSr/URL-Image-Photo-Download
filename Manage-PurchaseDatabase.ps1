<#
.SYNOPSIS
    Interactive database manager for purchase tracking data.

.DESCRIPTION
    Provides an interactive menu to view, edit, and update purchase tracking data
    including manual price overrides, bulk updates, and data maintenance.

.PARAMETER DatabasePath
    Path to the SQLite database file.

.PARAMETER DirectSQL
    Switch to enable direct SQL command mode (advanced users).

.EXAMPLE
    .\Manage-PurchaseDatabase.ps1 -DatabasePath ".\purchasesdata.db"

.EXAMPLE
    .\Manage-PurchaseDatabase.ps1 -DatabasePath ".\purchasesdata.db" -DirectSQL

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-10-07
    Version: 1.0.0
    Change Date: 
    Change Purpose: 
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
    [string]$DatabasePath,
    
    [Parameter(Mandatory=$false)]
    [switch]$DirectSQL
)

Import-Module PSSQLite -ErrorAction Stop

#region Helper Functions

function Show-Items {
    param(
        [string]$Filter = "",
        [int]$Limit = 50
    )
    
    $query = @"
        SELECT 
            ItemID,
            Description,
            Quantity,
            ROUND(PerItemCost, 2) as Cost,
            ROUND(CurrentMarketAvg, 2) as MarketAvg,
            ROUND(SuggestedPrice, 2) as OurPrice,
            ROUND(MinAcceptablePrice, 2) as MinPrice,
            Category
        FROM Items 
        WHERE IsActive = 1
        $(if ($Filter) { "AND Description LIKE '%$Filter%'" })
        ORDER BY ItemID
        LIMIT $Limit
"@
    
    $items = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
    
    if ($items) {
        Write-Host "`n=== Current Items ===" -ForegroundColor Cyan
        $items | Format-Table -AutoSize
        return $items
    }
    else {
        Write-Host "No items found" -ForegroundColor Yellow
        return $null
    }
}

function Update-ItemPrice {
    param(
        [int]$ItemID,
        [decimal]$NewPrice,
        [string]$PriceField = "SuggestedPrice"
    )
    
    $validFields = @('SuggestedPrice', 'MinAcceptablePrice', 'CurrentMSRP', 'CurrentMarketAvg', 'PerItemCost')
    
    if ($PriceField -notin $validFields) {
        Write-Host "Invalid field. Valid fields: $($validFields -join ', ')" -ForegroundColor Red
        return
    }
    
    try {
        $updateQuery = @"
            UPDATE Items 
            SET $PriceField = @Price,
                ModifiedDate = CURRENT_TIMESTAMP
            WHERE ItemID = @ItemID
"@
        $params = @{
            Price = $NewPrice
            ItemID = $ItemID
        }
        
        Invoke-SqliteQuery -DataSource $DatabasePath -Query $updateQuery -SqlParameters $params
        
        # Get updated item to show
        $checkQuery = "SELECT Description, $PriceField FROM Items WHERE ItemID = @ItemID"
        $updated = Invoke-SqliteQuery -DataSource $DatabasePath -Query $checkQuery -SqlParameters @{ItemID = $ItemID}
        
        Write-Host "✅ Updated successfully!" -ForegroundColor Green
        Write-Host "   Item: $($updated.Description)" -ForegroundColor White
        Write-Host "   $PriceField = `$$NewPrice" -ForegroundColor Green
        
        # Log the manual update
        $logQuery = @"
            INSERT INTO PriceHistory (ItemID, Source, Price, Confidence, URL)
            VALUES (@ItemID, 'Manual Override', @Price, 100, '$PriceField updated by user')
"@
        Invoke-SqliteQuery -DataSource $DatabasePath -Query $logQuery -SqlParameters $params
    }
    catch {
        Write-Host "Error updating item: $_" -ForegroundColor Red
    }
}

function Update-BulkPrices {
    param(
        [string]$Condition,
        [decimal]$Adjustment,
        [string]$AdjustmentType = "Percentage" # Percentage or Fixed
    )
    
    try {
        if ($AdjustmentType -eq "Percentage") {
            $updateQuery = @"
                UPDATE Items 
                SET SuggestedPrice = ROUND(SuggestedPrice * (1 + @Adjustment / 100), 2),
                    ModifiedDate = CURRENT_TIMESTAMP
                WHERE IsActive = 1 AND SuggestedPrice > 0
                $(if ($Condition) { "AND $Condition" })
"@
        }
        else {
            $updateQuery = @"
                UPDATE Items 
                SET SuggestedPrice = ROUND(SuggestedPrice + @Adjustment, 2),
                    ModifiedDate = CURRENT_TIMESTAMP
                WHERE IsActive = 1 AND SuggestedPrice > 0
                $(if ($Condition) { "AND $Condition" })
"@
        }
        
        $params = @{ Adjustment = $Adjustment }
        
        # Show what will be affected
        $previewQuery = @"
            SELECT COUNT(*) as Count 
            FROM Items 
            WHERE IsActive = 1 AND SuggestedPrice > 0
            $(if ($Condition) { "AND $Condition" })
"@
        $affected = Invoke-SqliteQuery -DataSource $DatabasePath -Query $previewQuery
        
        Write-Host "This will affect $($affected.Count) items" -ForegroundColor Yellow
        $confirm = Read-Host "Continue? (Y/N)"
        
        if ($confirm -eq 'Y') {
            Invoke-SqliteQuery -DataSource $DatabasePath -Query $updateQuery -SqlParameters $params
            Write-Host "✅ Bulk update completed!" -ForegroundColor Green
        }
        else {
            Write-Host "Cancelled" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error in bulk update: $_" -ForegroundColor Red
    }
}

function Edit-ItemDetails {
    param([int]$ItemID)
    
    # Get current item data
    $query = "SELECT * FROM Items WHERE ItemID = @ItemID"
    $item = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query -SqlParameters @{ItemID = $ItemID}
    
    if (-not $item) {
        Write-Host "Item not found" -ForegroundColor Red
        return
    }
    
    Write-Host "`n=== Editing Item ID: $ItemID ===" -ForegroundColor Cyan
    Write-Host "Current Description: $($item.Description)" -ForegroundColor Yellow
    
    $fields = @(
        @{Name='Description'; Type='String'; Current=$item.Description},
        @{Name='Category'; Type='String'; Current=$item.Category},
        @{Name='Quantity'; Type='Int'; Current=$item.Quantity},
        @{Name='PerItemCost'; Type='Decimal'; Current=$item.PerItemCost},
        @{Name='SuggestedPrice'; Type='Decimal'; Current=$item.SuggestedPrice},
        @{Name='MinAcceptablePrice'; Type='Decimal'; Current=$item.MinAcceptablePrice},
        @{Name='Notes'; Type='String'; Current=$item.Notes}
    )
    
    Write-Host "`nEditable Fields:" -ForegroundColor White
    for ($i = 0; $i -lt $fields.Count; $i++) {
        $current = if ($fields[$i].Current) { $fields[$i].Current } else { "(empty)" }
        Write-Host "$($i+1). $($fields[$i].Name): $current"
    }
    Write-Host "0. Done editing"
    
    do {
        $choice = Read-Host "`nSelect field to edit (0-$($fields.Count))"
        
        if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $fields.Count) {
            $fieldIndex = [int]$choice - 1
            $field = $fields[$fieldIndex]
            
            Write-Host "Current $($field.Name): $($field.Current)" -ForegroundColor Yellow
            $newValue = Read-Host "Enter new value (or press Enter to skip)"
            
            if ($newValue) {
                # Type conversion
                $convertedValue = switch ($field.Type) {
                    'Int' { [int]$newValue }
                    'Decimal' { [decimal]$newValue }
                    default { $newValue }
                }
                
                # Update database
                $updateQuery = "UPDATE Items SET $($field.Name) = @Value, ModifiedDate = CURRENT_TIMESTAMP WHERE ItemID = @ItemID"
                $params = @{
                    Value = $convertedValue
                    ItemID = $ItemID
                }
                
                try {
                    Invoke-SqliteQuery -DataSource $DatabasePath -Query $updateQuery -SqlParameters $params
                    Write-Host "✅ Updated $($field.Name) successfully!" -ForegroundColor Green
                    $fields[$fieldIndex].Current = $convertedValue
                }
                catch {
                    Write-Host "Error updating: $_" -ForegroundColor Red
                }
            }
        }
    } while ($choice -ne '0')
}

function Export-ForExcel {
    $exportPath = "PurchaseData_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $query = @"
        SELECT 
            ItemID,
            Lot,
            Description,
            Category,
            Quantity,
            PerItemCost,
            TotalCost,
            CurrentMSRP,
            CurrentMarketAvg,
            SuggestedPrice as 'Our Asking Price',
            MinAcceptablePrice as 'Minimum Accept Price',
            SoldPrice,
            Notes,
            Address,
            Location
        FROM Items 
        WHERE IsActive = 1
        ORDER BY ItemID
"@
    
    $data = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
    $data | Export-Csv -Path $exportPath -NoTypeInformation
    
    Write-Host "✅ Exported to: $exportPath" -ForegroundColor Green
    Write-Host "You can now edit in Excel and re-import" -ForegroundColor Cyan
    return $exportPath
}

function Import-FromExcel {
    param([string]$CSVPath)
    
    if (-not (Test-Path $CSVPath)) {
        Write-Host "File not found: $CSVPath" -ForegroundColor Red
        return
    }
    
    $data = Import-Csv -Path $CSVPath
    $updated = 0
    
    foreach ($row in $data) {
        if ($row.ItemID) {
            try {
                $updateQuery = @"
                    UPDATE Items SET
                        Description = @Description,
                        Category = @Category,
                        Quantity = @Quantity,
                        PerItemCost = @PerItemCost,
                        SuggestedPrice = @OurPrice,
                        MinAcceptablePrice = @MinPrice,
                        Notes = @Notes,
                        ModifiedDate = CURRENT_TIMESTAMP
                    WHERE ItemID = @ItemID
"@
                $params = @{
                    ItemID = [int]$row.ItemID
                    Description = $row.Description
                    Category = $row.Category
                    Quantity = if ($row.Quantity) { [int]$row.Quantity } else { 1 }
                    PerItemCost = if ($row.PerItemCost) { [decimal]($row.PerItemCost -replace '[$,]','')} else { 0 }
                    OurPrice = if ($row.'Our Asking Price') { [decimal]($row.'Our Asking Price' -replace '[$,]','')} else { 0 }
                    MinPrice = if ($row.'Minimum Accept Price') { [decimal]($row.'Minimum Accept Price' -replace '[$,]','')} else { 0 }
                    Notes = $row.Notes
                }
                
                Invoke-SqliteQuery -DataSource $DatabasePath -Query $updateQuery -SqlParameters $params
                $updated++
            }
            catch {
                Write-Host "Error updating ItemID $($row.ItemID): $_" -ForegroundColor Red
            }
        }
    }
    
    Write-Host "✅ Updated $updated items from Excel" -ForegroundColor Green
}

#endregion

#region Main Menu

function Show-Menu {
    Clear-Host
    Write-Host @"
╔══════════════════════════════════════════════════════════════╗
║            PURCHASE DATABASE MANAGER                          ║
╚══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan
    
    Write-Host "`nDatabase: $DatabasePath" -ForegroundColor Yellow
    
    # Quick stats
    $stats = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
        SELECT 
            COUNT(*) as Items,
            ROUND(SUM(PerItemCost * Quantity), 2) as Investment,
            ROUND(SUM(SuggestedPrice * Quantity), 2) as Revenue
        FROM Items WHERE IsActive = 1
"@
    
    Write-Host "Items: $($stats.Items) | Investment: `$$($stats.Investment) | Potential: `$$($stats.Revenue)" -ForegroundColor Gray
    
    Write-Host "`n=== MAIN MENU ===" -ForegroundColor Green
    Write-Host "1. View all items"
    Write-Host "2. Search items"
    Write-Host "3. Edit single item prices"
    Write-Host "4. Edit all item details"
    Write-Host "5. Bulk price adjustment"
    Write-Host "6. Export to Excel (CSV)"
    Write-Host "7. Import from Excel (CSV)"
    Write-Host "8. Mark item as sold"
    Write-Host "9. Add new item manually"
    Write-Host "10. View price history"
    Write-Host "11. Delete item"
    Write-Host "12. Run custom SQL query" -ForegroundColor Yellow
    Write-Host "0. Exit"
    Write-Host ""
}

#endregion

# Direct SQL Mode
if ($DirectSQL) {
    Write-Host "=== DIRECT SQL MODE ===" -ForegroundColor Yellow
    Write-Host "Type 'exit' to quit, 'help' for common queries" -ForegroundColor Cyan
    Write-Host ""
    
    while ($true) {
        $query = Read-Host "SQL"
        
        if ($query -eq 'exit') { break }
        
        if ($query -eq 'help') {
            Write-Host @"
            
Common SQL Queries:
-------------------
View all items:
  SELECT * FROM Items WHERE IsActive = 1

Update single price:
  UPDATE Items SET SuggestedPrice = 150 WHERE ItemID = 1

Bulk price increase by 10%:
  UPDATE Items SET SuggestedPrice = SuggestedPrice * 1.10 WHERE Category = 'Electronics'

View items needing prices:
  SELECT * FROM Items WHERE SuggestedPrice IS NULL OR SuggestedPrice = 0

View price history:
  SELECT * FROM PriceHistory ORDER BY CaptureDate DESC LIMIT 20

"@ -ForegroundColor Cyan
            continue
        }
        
        try {
            if ($query -match '^SELECT|^PRAGMA') {
                $result = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
                $result | Format-Table -AutoSize
            }
            else {
                Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
                Write-Host "✅ Query executed successfully" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Error: $_" -ForegroundColor Red
        }
    }
    exit
}

# Interactive Menu Mode
do {
    Show-Menu
    $choice = Read-Host "Select option"
    
    switch ($choice) {
        '1' { # View all items
            Show-Items
            Read-Host "`nPress Enter to continue"
        }
        
        '2' { # Search items
            $search = Read-Host "Enter search term"
            Show-Items -Filter $search
            Read-Host "`nPress Enter to continue"
        }
        
        '3' { # Edit single item prices
            $items = Show-Items -Limit 100
            if ($items) {
                Write-Host "`n=== PRICE EDITOR ===" -ForegroundColor Cyan
                $itemId = Read-Host "Enter ItemID to edit"
                
                if ($itemId -match '^\d+$') {
                    Write-Host "`nWhich price to update?" -ForegroundColor Yellow
                    Write-Host "1. Our Asking Price (SuggestedPrice)"
                    Write-Host "2. Minimum Accept Price (MinAcceptablePrice)"
                    Write-Host "3. Market Average (CurrentMarketAvg)"
                    Write-Host "4. MSRP (CurrentMSRP)"
                    Write-Host "5. Per Item Cost"
                    
                    $priceChoice = Read-Host "Select (1-5)"
                    
                    $fieldMap = @{
                        '1' = 'SuggestedPrice'
                        '2' = 'MinAcceptablePrice'
                        '3' = 'CurrentMarketAvg'
                        '4' = 'CurrentMSRP'
                        '5' = 'PerItemCost'
                    }
                    
                    if ($fieldMap.ContainsKey($priceChoice)) {
                        $newPrice = Read-Host "Enter new price (numbers only, no $)"
                        if ($newPrice -match '^\d+\.?\d*$') {
                            Update-ItemPrice -ItemID $itemId -NewPrice ([decimal]$newPrice) -PriceField $fieldMap[$priceChoice]
                        }
                        else {
                            Write-Host "Invalid price format" -ForegroundColor Red
                        }
                    }
                }
            }
            Read-Host "`nPress Enter to continue"
        }
        
        '4' { # Edit all item details
            $items = Show-Items -Limit 100
            if ($items) {
                $itemId = Read-Host "`nEnter ItemID to edit"
                if ($itemId -match '^\d+$') {
                    Edit-ItemDetails -ItemID ([int]$itemId)
                }
            }
            Read-Host "`nPress Enter to continue"
        }
        
        '5' { # Bulk price adjustment
            Write-Host "`n=== BULK PRICE ADJUSTMENT ===" -ForegroundColor Cyan
            Write-Host "1. Increase all prices by percentage"
            Write-Host "2. Decrease all prices by percentage"
            Write-Host "3. Increase prices for specific category"
            Write-Host "4. Round all prices to nearest $5"
            Write-Host "5. Set minimum profit margin"
            
            $bulkChoice = Read-Host "Select option"
            
            switch ($bulkChoice) {
                '1' {
                    $percent = Read-Host "Enter percentage increase (e.g., 10 for 10%)"
                    Update-BulkPrices -Adjustment ([decimal]$percent) -AdjustmentType "Percentage"
                }
                '2' {
                    $percent = Read-Host "Enter percentage decrease (e.g., 10 for 10%)"
                    Update-BulkPrices -Adjustment ([decimal]-$percent) -AdjustmentType "Percentage"
                }
                '3' {
                    $category = Read-Host "Enter category"
                    $percent = Read-Host "Enter percentage adjustment"
                    Update-BulkPrices -Condition "Category = '$category'" -Adjustment ([decimal]$percent) -AdjustmentType "Percentage"
                }
                '4' {
                    $roundQuery = @"
                        UPDATE Items 
                        SET SuggestedPrice = ROUND(SuggestedPrice / 5.0) * 5,
                            ModifiedDate = CURRENT_TIMESTAMP
                        WHERE IsActive = 1 AND SuggestedPrice > 0
"@
                    Invoke-SqliteQuery -DataSource $DatabasePath -Query $roundQuery
                    Write-Host "✅ Prices rounded to nearest `$5" -ForegroundColor Green
                }
                '5' {
                    $margin = Read-Host "Enter minimum profit margin % (e.g., 200 for 200%)"
                    $marginQuery = @"
                        UPDATE Items 
                        SET SuggestedPrice = ROUND(PerItemCost * (1 + @Margin / 100), 2)
                        WHERE IsActive = 1 
                        AND PerItemCost > 0
                        AND (SuggestedPrice < PerItemCost * (1 + @Margin / 100) OR SuggestedPrice IS NULL)
"@
                    Invoke-SqliteQuery -DataSource $DatabasePath -Query $marginQuery -SqlParameters @{Margin = [decimal]$margin}
                    Write-Host "✅ Minimum profit margin set" -ForegroundColor Green
                }
            }
            Read-Host "`nPress Enter to continue"
        }
        
        '6' { # Export to Excel
            $exportFile = Export-ForExcel
            Write-Host "`nYou can now:"
            Write-Host "1. Open in Excel: $exportFile"
            Write-Host "2. Edit prices in the 'Our Asking Price' column"
            Write-Host "3. Save the file"
            Write-Host "4. Use option 7 to import changes"
            Read-Host "`nPress Enter to continue"
        }
        
        '7' { # Import from Excel
            $importPath = Read-Host "Enter path to CSV file"
            Import-FromExcel -CSVPath $importPath
            Read-Host "`nPress Enter to continue"
        }
        
        '8' { # Mark as sold
            Show-Items
            $itemId = Read-Host "`nEnter ItemID to mark as sold"
            if ($itemId -match '^\d+$') {
                $soldPrice = Read-Host "Enter sold price"
                $soldDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                
                $soldQuery = @"
                    UPDATE Items 
                    SET SoldPrice = @Price,
                        SoldDate = @Date,
                        IsActive = 0,
                        ModifiedDate = CURRENT_TIMESTAMP
                    WHERE ItemID = @ItemID
"@
                $params = @{
                    Price = [decimal]$soldPrice
                    Date = $soldDate
                    ItemID = [int]$itemId
                }
                
                Invoke-SqliteQuery -DataSource $DatabasePath -Query $soldQuery -SqlParameters $params
                Write-Host "✅ Item marked as sold for `$$soldPrice" -ForegroundColor Green
            }
            Read-Host "`nPress Enter to continue"
        }
        
        '9' { # Add new item
            Write-Host "`n=== ADD NEW ITEM ===" -ForegroundColor Cyan
            $description = Read-Host "Description"
            $quantity = Read-Host "Quantity (default 1)"
            $cost = Read-Host "Per Item Cost"
            $category = Read-Host "Category (optional)"
            
            $insertQuery = @"
                INSERT INTO Items (Description, Quantity, PerItemCost, Category, IsActive)
                VALUES (@Desc, @Qty, @Cost, @Cat, 1)
"@
            $params = @{
                Desc = $description
                Qty = if ($quantity) { [int]$quantity } else { 1 }
                Cost = [decimal]$cost
                Cat = if ($category) { $category } else { 'General' }
            }
            
            Invoke-SqliteQuery -DataSource $DatabasePath -Query $insertQuery -SqlParameters $params
            Write-Host "✅ Item added successfully" -ForegroundColor Green
            Read-Host "`nPress Enter to continue"
        }
        
        '10' { # View price history
            $historyQuery = @"
                SELECT 
                    p.PriceID,
                    i.Description,
                    p.Source,
                    ROUND(p.Price, 2) as Price,
                    p.Confidence,
                    p.CaptureDate
                FROM PriceHistory p
                JOIN Items i ON p.ItemID = i.ItemID
                ORDER BY p.CaptureDate DESC
                LIMIT 50
"@
            $history = Invoke-SqliteQuery -DataSource $DatabasePath -Query $historyQuery
            $history | Format-Table -AutoSize
            Read-Host "`nPress Enter to continue"
        }
        
        '11' { # Delete item
            Show-Items
            $itemId = Read-Host "`nEnter ItemID to delete (or mark inactive)"
            if ($itemId -match '^\d+$') {
                Write-Host "1. Mark as inactive (can be restored)"
                Write-Host "2. Permanently delete"
                $delChoice = Read-Host "Select option"
                
                if ($delChoice -eq '1') {
                    $query = "UPDATE Items SET IsActive = 0 WHERE ItemID = @ItemID"
                    Invoke-SqliteQuery -DataSource $DatabasePath -Query $query -SqlParameters @{ItemID = [int]$itemId}
                    Write-Host "✅ Item marked as inactive" -ForegroundColor Green
                }
                elseif ($delChoice -eq '2') {
                    $confirm = Read-Host "Are you sure? This cannot be undone! (YES to confirm)"
                    if ($confirm -eq 'YES') {
                        $query = "DELETE FROM Items WHERE ItemID = @ItemID"
                        Invoke-SqliteQuery -DataSource $DatabasePath -Query $query -SqlParameters @{ItemID = [int]$itemId}
                        Write-Host "✅ Item permanently deleted" -ForegroundColor Green
                    }
                }
            }
            Read-Host "`nPress Enter to continue"
        }
        
        '12' { # Custom SQL
            Write-Host "`n=== CUSTOM SQL QUERY ===" -ForegroundColor Yellow
            Write-Host "Enter your SQL query (be careful with UPDATE/DELETE!)" -ForegroundColor Red
            $query = Read-Host "SQL"
            
            try {
                if ($query -match '^SELECT|^PRAGMA') {
                    $result = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
                    $result | Format-Table -AutoSize
                }
                else {
                    Write-Host "This is a modification query. Continue? (Y/N)" -ForegroundColor Yellow
                    $confirm = Read-Host
                    if ($confirm -eq 'Y') {
                        Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
                        Write-Host "✅ Query executed" -ForegroundColor Green
                    }
                }
            }
            catch {
                Write-Host "Error: $_" -ForegroundColor Red
            }
            Read-Host "`nPress Enter to continue"
        }
    }
} while ($choice -ne '0')

Write-Host "`nGoodbye!" -ForegroundColor Cyan