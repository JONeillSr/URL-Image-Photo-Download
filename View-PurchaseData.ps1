<#
.SYNOPSIS
    View and analyze purchase tracking data.

.DESCRIPTION
    Displays current inventory, pricing recommendations, and profit analysis.

.PARAMETER DatabasePath
    Path to the SQLite database file.

.PARAMETER Export
    Switch to export data to Excel-compatible CSV.

.EXAMPLE
    .\View-PurchaseData.ps1 -DatabasePath ".\purchasesdata.db"

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
    [switch]$Export
)

Import-Module PSSQLite -ErrorAction Stop

Clear-Host

Write-Host @"
╔════════════════════════════════════════════════════════════════╗
║          PURCHASE TRACKING - PRICING ANALYSIS REPORT          ║
╚════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

Write-Host "`nDatabase: $DatabasePath" -ForegroundColor Yellow
Write-Host "Report Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow

# Get summary statistics
$stats = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
    SELECT 
        COUNT(*) as TotalItems,
        SUM(Quantity) as TotalUnits,
        ROUND(SUM(PerItemCost * Quantity), 2) as TotalInvestment,
        ROUND(SUM(SuggestedPrice * Quantity), 2) as PotentialRevenue,
        ROUND(SUM((SuggestedPrice - PerItemCost) * Quantity), 2) as PotentialProfit
    FROM Items 
    WHERE IsActive = 1
"@

Write-Host "`n📊 PORTFOLIO SUMMARY" -ForegroundColor Green
Write-Host "══════════════════════════════════════════" -ForegroundColor DarkGray
Write-Host "Total Items:        $($stats.TotalItems)" -ForegroundColor White
Write-Host "Total Units:        $($stats.TotalUnits)" -ForegroundColor White
Write-Host "Total Investment:   `$$($stats.TotalInvestment)" -ForegroundColor Yellow
Write-Host "Potential Revenue:  `$$($stats.PotentialRevenue)" -ForegroundColor Green
Write-Host "Potential Profit:   `$$($stats.PotentialProfit)" -ForegroundColor Green -BackgroundColor DarkGreen

if ($stats.TotalInvestment -gt 0) {
    $roi = [Math]::Round(($stats.PotentialProfit / $stats.TotalInvestment) * 100, 1)
    Write-Host "ROI:                $roi%" -ForegroundColor Cyan
}

# Get items with best margins
Write-Host "`n💰 TOP PROFIT OPPORTUNITIES" -ForegroundColor Green
Write-Host "══════════════════════════════════════════" -ForegroundColor DarkGray

$topItems = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
    SELECT 
        Description,
        Quantity,
        ROUND(PerItemCost, 2) as Cost,
        ROUND(CurrentMarketAvg, 2) as MarketPrice,
        ROUND(SuggestedPrice, 2) as OurPrice,
        ROUND((SuggestedPrice - PerItemCost), 2) as ProfitPerItem,
        ROUND(((SuggestedPrice / PerItemCost - 1) * 100), 1) as Margin
    FROM Items 
    WHERE IsActive = 1 AND SuggestedPrice > 0 AND PerItemCost > 0
    ORDER BY Margin DESC
    LIMIT 5
"@

foreach ($item in $topItems) {
    Write-Host "`n$($item.Description)" -ForegroundColor White
    Write-Host "  Qty: $($item.Quantity) | Cost: `$$($item.Cost) | Market: `$$($item.MarketPrice) | " -NoNewline
    Write-Host "Our Price: `$$($item.OurPrice)" -ForegroundColor Green -NoNewline
    Write-Host " | Profit: `$$($item.ProfitPerItem) (" -NoNewline
    Write-Host "$($item.Margin)%" -ForegroundColor Cyan -NoNewline
    Write-Host ")"
}

# Get items needing price updates
Write-Host "`n⚠️  ITEMS NEEDING ATTENTION" -ForegroundColor Yellow
Write-Host "══════════════════════════════════════════" -ForegroundColor DarkGray

$needsUpdate = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
    SELECT Description, PerItemCost
    FROM Items 
    WHERE IsActive = 1 
    AND (CurrentMarketAvg IS NULL OR CurrentMarketAvg = 0)
"@

if ($needsUpdate) {
    foreach ($item in $needsUpdate) {
        Write-Host "  ❌ $($item.Description) - No market pricing" -ForegroundColor Yellow
    }
}
else {
    Write-Host "  ✅ All items have pricing data!" -ForegroundColor Green
}

# Recent price captures
Write-Host "`n📈 RECENT MARKET PRICES" -ForegroundColor Green
Write-Host "══════════════════════════════════════════" -ForegroundColor DarkGray

$recentPrices = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
    SELECT 
        i.Description,
        p.Source,
        ROUND(p.Price, 2) as Price,
        p.CaptureDate
    FROM PriceHistory p
    JOIN Items i ON p.ItemID = i.ItemID
    ORDER BY p.CaptureDate DESC
    LIMIT 10
"@

if ($recentPrices) {
    $recentPrices | Format-Table -AutoSize
}
else {
    Write-Host "  No price history recorded yet" -ForegroundColor Yellow
}

# Export option
if ($Export) {
    $exportPath = "PurchaseAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $exportData = Invoke-SqliteQuery -DataSource $DatabasePath -Query @"
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
            SuggestedPrice,
            MinAcceptablePrice,
            (SuggestedPrice - PerItemCost) as ProfitPerItem,
            ROUND(((SuggestedPrice / PerItemCost - 1) * 100), 1) as MarginPercent,
            Address,
            Location,
            Notes
        FROM Items 
        WHERE IsActive = 1
        ORDER BY ItemID
"@
    
    $exportData | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "`n✅ Data exported to: $exportPath" -ForegroundColor Green
}

Write-Host "`n" -NoNewline
Write-Host "══════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "         END OF REPORT - Azure Innovators Analytics" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════════════" -ForegroundColor Cyan