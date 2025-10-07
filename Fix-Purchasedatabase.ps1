<#
.SYNOPSIS
    Updates existing Purchase Tracking database schema and fixes issues.

.DESCRIPTION
    This script updates an existing database to add missing columns and fix schema issues.

.PARAMETER DatabasePath
    Path to the existing SQLite database file.

.PARAMETER ShowData
    Switch to display current data after update.

.EXAMPLE
    .\Fix-PurchaseDatabase.ps1 -DatabasePath ".\purchasesdata.db"

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
    [switch]$ShowData
)

# Import SQLite module
try {
    Import-Module PSSQLite -ErrorAction Stop
}
catch {
    Write-Error "PSSQLite module not found. Please install: Install-Module PSSQLite -Scope CurrentUser"
    exit 1
}

Write-Host "`n=== Database Schema Update Tool ===" -ForegroundColor Cyan
Write-Host "Database: $DatabasePath" -ForegroundColor Yellow

# Check current schema
Write-Host "`nChecking current schema..." -ForegroundColor Yellow

try {
    # Check if PriceHistory table exists
    $tableCheck = "SELECT name FROM sqlite_master WHERE type='table' AND name='PriceHistory'"
    $priceTable = Invoke-SqliteQuery -DataSource $DatabasePath -Query $tableCheck
    
    if ($priceTable) {
        Write-Host "PriceHistory table exists" -ForegroundColor Green
        
        # Check columns
        $columnCheck = "PRAGMA table_info(PriceHistory)"
        $columns = Invoke-SqliteQuery -DataSource $DatabasePath -Query $columnCheck
        
        Write-Host "Current columns in PriceHistory:" -ForegroundColor Yellow
        $columns | ForEach-Object { Write-Host "  - $($_.name)" -ForegroundColor Gray }
        
        # Check if Confidence column exists
        $hasConfidence = $columns | Where-Object { $_.name -eq 'Confidence' }
        
        if (-not $hasConfidence) {
            Write-Host "`nAdding Confidence column..." -ForegroundColor Yellow
            
            $addColumn = "ALTER TABLE PriceHistory ADD COLUMN Confidence REAL DEFAULT 75"
            Invoke-SqliteQuery -DataSource $DatabasePath -Query $addColumn
            
            Write-Host "Confidence column added successfully!" -ForegroundColor Green
        }
        else {
            Write-Host "Confidence column already exists" -ForegroundColor Green
        }
    }
    else {
        Write-Host "PriceHistory table doesn't exist, creating it..." -ForegroundColor Yellow
        
        $createTable = @"
CREATE TABLE PriceHistory (
    PriceID INTEGER PRIMARY KEY AUTOINCREMENT,
    ItemID INTEGER NOT NULL,
    Source TEXT NOT NULL,
    URL TEXT,
    Price REAL,
    Confidence REAL DEFAULT 75,
    CaptureDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ItemID) REFERENCES Items(ItemID)
)
"@
        Invoke-SqliteQuery -DataSource $DatabasePath -Query $createTable
        Write-Host "PriceHistory table created!" -ForegroundColor Green
    }
    
    # Fix ItemID = 0 issues
    Write-Host "`nChecking for ItemID issues..." -ForegroundColor Yellow
    
    $zeroIdCheck = "SELECT COUNT(*) as Count FROM Items WHERE ItemID = 0"
    $zeroCount = Invoke-SqliteQuery -DataSource $DatabasePath -Query $zeroIdCheck
    
    if ($zeroCount.Count -gt 0) {
        Write-Host "Found $($zeroCount.Count) items with ItemID = 0" -ForegroundColor Yellow
        Write-Host "These will be fixed automatically on next run" -ForegroundColor Yellow
    }
    
    # Show summary
    Write-Host "`n=== Database Summary ===" -ForegroundColor Cyan
    
    $itemCount = Invoke-SqliteQuery -DataSource $DatabasePath -Query "SELECT COUNT(*) as Count FROM Items"
    $priceCount = Invoke-SqliteQuery -DataSource $DatabasePath -Query "SELECT COUNT(*) as Count FROM PriceHistory"
    
    Write-Host "Total Items: $($itemCount.Count)" -ForegroundColor White
    Write-Host "Total Price Records: $($priceCount.Count)" -ForegroundColor White
    
    if ($ShowData) {
        Write-Host "`n=== Current Items ===" -ForegroundColor Cyan
        
        $items = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
            SELECT 
                ItemID,
                Description,
                PerItemCost,
                CurrentMarketAvg,
                SuggestedPrice,
                CASE 
                    WHEN SuggestedPrice > 0 AND PerItemCost > 0 
                    THEN ROUND((SuggestedPrice / PerItemCost - 1) * 100, 1)
                    ELSE 0 
                END as ProfitMargin
            FROM Items 
            WHERE IsActive = 1
            ORDER BY ItemID
"@
        
        $items | Format-Table -AutoSize
        
        Write-Host "`n=== Recent Price History ===" -ForegroundColor Cyan
        
        $prices = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
            SELECT 
                p.PriceID,
                i.Description,
                p.Source,
                p.Price,
                p.Confidence,
                p.CaptureDate
            FROM PriceHistory p
            JOIN Items i ON p.ItemID = i.ItemID
            ORDER BY p.CaptureDate DESC
            LIMIT 10
"@
        
        if ($prices) {
            $prices | Format-Table -AutoSize
        }
        else {
            Write-Host "No price history yet" -ForegroundColor Yellow
        }
    }
    
    Write-Host "`n✓ Database is ready for use!" -ForegroundColor Green
    Write-Host "You can now run the main script without errors." -ForegroundColor Cyan
}
catch {
    Write-Host "`nError updating database: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}