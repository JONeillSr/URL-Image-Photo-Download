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
    Create Date: 2025-10-07
    Version: 2.0.0
    Change Date: 
    Change Purpose: 
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

function ConvertTo-Hashtable {
    param($InputObject)
    
    $hash = @{}
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $InputObject.PSObject.Properties | ForEach-Object {
            $hash[$_.Name] = $_.Value
        }
    }
    elseif ($InputObject -is [System.Collections.IDictionary]) {
        return $InputObject
    }
    return $hash
}

function Normalize-ItemData {
    param($Item)
    
    $normalized = @{}
    
    # Map CSV column names to database column names
    $columnMapping = @{
        'Lot' = 'Lot'
        'Description' = 'Description'
        'Address' = 'Address'
        'Plant' = 'Plant'
        'Location' = 'Location'
        'Quantity' = 'Quantity'
        'Bid' = 'Bid'
        'Sale Price' = 'SalePrice'
        'Premium' = 'Premium'
        'Tax' = 'Tax'
        'Rigging Fee' = 'RiggingFee'
        'Freight Cost' = 'FreightCost'
        'Other Costs' = 'OtherCosts'
        'Total' = 'TotalCost'
        'Per Item Cost' = 'PerItemCost'
        'Category' = 'Category'
        'Notes' = 'Notes'
        'Photos' = 'Photos'
        'MSRP' = 'CurrentMSRP'
        'Average On-Line' = 'CurrentMarketAvg'
        'Our Asking Price' = 'SuggestedPrice'
        'Sold Price' = 'SoldPrice'
    }
    
    # Convert PSCustomObject to hashtable first
    $itemHash = ConvertTo-Hashtable -InputObject $Item
    
    foreach ($csvColumn in $columnMapping.Keys) {
        $dbColumn = $columnMapping[$csvColumn]
        if ($itemHash.ContainsKey($csvColumn)) {
            $value = $itemHash[$csvColumn]
            
            # Clean currency values
            if ($value -is [string] -and $value -match '^\$[\d,]+\.?\d*$') {
                $value = $value -replace '[$,]', ''
            }
            
            # Convert empty strings to null
            if ($value -eq '') {
                $value = $null
            }
            
            # Convert to appropriate type
            if ($null -ne $value) {
                try {
                    if ($dbColumn -match 'Price|Cost|Fee|Premium|Tax|Total|MSRP|Avg') {
                        $value = [decimal]$value
                    }
                    elseif ($dbColumn -match 'Quantity|Lot') {
                        if ($value -match '^\d+$') {
                            $value = [int]$value
                        }
                    }
                    elseif ($dbColumn -eq 'Plant') {
                        # Plant might be string or number
                        if ($value -match '^\d+$') {
                            $value = [int]$value
                        }
                    }
                }
                catch {
                    Write-Warning "Could not convert '$csvColumn' value '$value'"
                }
            }
            
            $normalized[$dbColumn] = $value
        }
    }
    
    return $normalized
}

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

#region Database Functions

function Initialize-Database {
    param([string]$DbPath)
    
    Write-Host "Initializing database at: $DbPath" -ForegroundColor Cyan
    
    $schema = @"
CREATE TABLE IF NOT EXISTS Items (
    ItemID INTEGER PRIMARY KEY AUTOINCREMENT,
    Lot INTEGER,
    Description TEXT NOT NULL,
    Category TEXT,
    Address TEXT,
    Plant TEXT,
    Location TEXT,
    Quantity INTEGER DEFAULT 1,
    Bid REAL,
    SalePrice REAL,
    Premium REAL,
    Tax REAL,
    RiggingFee REAL,
    FreightCost REAL,
    OtherCosts REAL,
    TotalCost REAL,
    PerItemCost REAL,
    CurrentMSRP REAL,
    CurrentMarketAvg REAL,
    SuggestedPrice REAL,
    MinAcceptablePrice REAL,
    SoldPrice REAL,
    Notes TEXT,
    Photos TEXT,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    IsActive BOOLEAN DEFAULT 1
);

CREATE TABLE IF NOT EXISTS PriceHistory (
    PriceID INTEGER PRIMARY KEY AUTOINCREMENT,
    ItemID INTEGER NOT NULL,
    Source TEXT NOT NULL,
    URL TEXT,
    Price REAL,
    Confidence REAL,
    CaptureDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ItemID) REFERENCES Items(ItemID)
);

CREATE INDEX IF NOT EXISTS idx_items_description ON Items(Description);
CREATE INDEX IF NOT EXISTS idx_price_history_item ON PriceHistory(ItemID);
"@
    
    try {
        Invoke-SqliteQuery -DataSource $DbPath -Query $schema
        Write-Host "Database initialized successfully!" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to initialize database: $_"
        throw
    }
}

#endregion

#region Web Scraping Functions

function Get-WebContent {
    param(
        [string]$URL,
        [int]$TimeoutSeconds = 30
    )
    
    $userAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    
    try {
        $response = Invoke-WebRequest -Uri $URL `
                                     -UserAgent $userAgent `
                                     -TimeoutSec $TimeoutSeconds `
                                     -UseBasicParsing `
                                     -ErrorAction Stop
        
        return @{
            Success = $true
            Content = $response.Content
            StatusCode = $response.StatusCode
        }
    }
    catch {
        return @{
            Success = $false
            Content = $null
            Error = $_.Exception.Message
        }
    }
}

function Search-ProductPrice {
    param(
        [string]$ProductDescription,
        [string]$Site
    )
    
    Write-Host "  Searching $Site..." -ForegroundColor Yellow
    
    # Build search URL
    $searchQuery = $ProductDescription -replace '\s+', '+'
    $searchUrl = switch ($Site) {
        'walmart.com' { "https://www.walmart.com/search?q=$searchQuery" }
        'amazon.com' { "https://www.amazon.com/s?k=$searchQuery" }
        'ebay.com' { "https://www.ebay.com/sch/i.html?_nkw=$searchQuery" }
        'lippert.com' { "https://shop.lippertcomponents.com/search?q=$searchQuery" }
        default { $null }
    }
    
    if (-not $searchUrl) {
        return $null
    }
    
    $webResult = Get-WebContent -URL $searchUrl
    
    if ($webResult.Success) {
        # Extract prices using regex patterns
        $pricePatterns = @(
            '\$([0-9]{1,4}\.?[0-9]{0,2})',
            '"price":\s*"?([0-9]+\.?[0-9]*)"?',
            'USD\s*([0-9]+\.?[0-9]*)'
        )
        
        $prices = @()
        foreach ($pattern in $pricePatterns) {
            $priceMatches = [regex]::Matches($webResult.Content, $pattern)
            foreach ($match in $priceMatches) {
                if ($match.Groups.Count -gt 1) {
                    $price = [double]$match.Groups[1].Value
                    if ($price -gt 10 -and $price -lt 10000) {
                        $prices += $price
                    }
                }
            }
        }
        
        if ($prices.Count -gt 0) {
            $avgPrice = ($prices | Measure-Object -Average).Average
            Write-Host "    Found price: `$$([Math]::Round($avgPrice, 2))" -ForegroundColor Green
            return $avgPrice
        }
    }
    
    Write-Host "    No prices found" -ForegroundColor Red
    return $null
}

#endregion

#region Processing Functions

function Invoke-PurchaseProcessing {
    param(
        [string]$CSVPath,
        [string]$DatabasePath,
        [switch]$UseAPI,
        [switch]$UpdateExisting
    )
    
    Write-Host "`n=== Purchase Processing Started ===" -ForegroundColor Cyan
    
    # Initialize database if needed
    if (-not (Test-Path $DatabasePath)) {
        Initialize-Database -DbPath $DatabasePath
    }
    
    # Load items
    $items = @()
    if ($CSVPath -and (Test-Path $CSVPath)) {
        Write-Host "Loading items from CSV..." -ForegroundColor Yellow
        $items = Import-Csv -Path $CSVPath
    }
    elseif ($UpdateExisting) {
        Write-Host "Loading items from database..." -ForegroundColor Yellow
        $query = "SELECT * FROM Items WHERE IsActive = 1"
        $items = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
    }
    
    if ($items.Count -eq 0) {
        Write-Warning "No items to process"
        return
    }
    
    $totalItems = $items.Count
    $currentItem = 0
    
    foreach ($item in $items) {
        $currentItem++
        Write-Host "`n[$currentItem/$totalItems] Processing: $($item.Description)" -ForegroundColor White
        
        # Normalize data if from CSV
        $normalizedItem = if ($CSVPath) {
            Normalize-ItemData -Item $item
        } else {
            ConvertTo-Hashtable -InputObject $item
        }
        
        # Insert or update in database
        $itemID = 0
        if ($CSVPath) {
            # Check if exists
            $checkQuery = "SELECT ItemID FROM Items WHERE Description = @Description"
            $checkParams = @{ Description = $normalizedItem.Description }
            $existing = Invoke-SqliteQuery -DataSource $DatabasePath `
                                          -Query $checkQuery `
                                          -SqlParameters $checkParams
            
            if ($existing) {
                $itemID = $existing.ItemID
                Write-Host "  Item already in database (ID: $itemID)" -ForegroundColor Yellow
            }
            else {
                # Insert new item
                $insertQuery = @"
                INSERT INTO Items (
                    Description, Lot, Quantity, PerItemCost, TotalCost,
                    Category, Address, Plant, Location
                ) VALUES (
                    @Description, @Lot, @Quantity, @PerItemCost, @TotalCost,
                    @Category, @Address, @Plant, @Location
                )
"@
                $insertParams = @{
                    Description = $normalizedItem.Description
                    Lot = if ($normalizedItem.Lot) { $normalizedItem.Lot } else { 0 }
                    Quantity = if ($normalizedItem.Quantity) { $normalizedItem.Quantity } else { 1 }
                    PerItemCost = if ($normalizedItem.PerItemCost) { $normalizedItem.PerItemCost } else { 0 }
                    TotalCost = if ($normalizedItem.TotalCost) { $normalizedItem.TotalCost } else { 0 }
                    Category = if ($normalizedItem.Category) { $normalizedItem.Category } else { 'General' }
                    Address = if ($normalizedItem.Address) { $normalizedItem.Address } else { '' }
                    Plant = if ($normalizedItem.Plant) { $normalizedItem.Plant } else { '' }
                    Location = if ($normalizedItem.Location) { $normalizedItem.Location } else { '' }
                }
                
                Invoke-SqliteQuery -DataSource $DatabasePath `
                                 -Query $insertQuery `
                                 -SqlParameters $insertParams
                
                # Get the ID
                $getIdQuery = "SELECT last_insert_rowid() as ItemID"
                $result = Invoke-SqliteQuery -DataSource $DatabasePath -Query $getIdQuery
                $itemID = $result.ItemID
                Write-Host "  Added to database (ID: $itemID)" -ForegroundColor Green
            }
        }
        else {
            $itemID = $normalizedItem.ItemID
        }
        
        # Search for prices
        $sites = @('walmart.com', 'amazon.com', 'ebay.com', 'lippert.com')
        $foundPrices = @()
        
        foreach ($site in $sites) {
            $price = Search-ProductPrice -ProductDescription $normalizedItem.Description -Site $site
            if ($price) {
                $foundPrices += $price
                
                # Store in price history
                $historyQuery = @"
                INSERT INTO PriceHistory (ItemID, Source, Price, Confidence)
                VALUES (@ItemID, @Source, @Price, @Confidence)
"@
                $historyParams = @{
                    ItemID = $itemID
                    Source = $site
                    Price = $price
                    Confidence = 75
                }
                
                Invoke-SqliteQuery -DataSource $DatabasePath `
                                 -Query $historyQuery `
                                 -SqlParameters $historyParams
            }
            
            Start-Sleep -Milliseconds 500
        }
        
        # Calculate recommendations
        if ($foundPrices.Count -gt 0) {
            $avgPrice = [Math]::Round(($foundPrices | Measure-Object -Average).Average, 2)
            $msrpPrice = [Math]::Round(($foundPrices | Measure-Object -Maximum).Maximum, 2)
            
            $cost = if ($normalizedItem.PerItemCost) { 
                [decimal]$normalizedItem.PerItemCost 
            } else { 
                1.0 
            }
            
            # Calculate suggested price
            $basePrice = $cost * 3.5
            $marketAdjusted = $avgPrice * 0.85
            $suggestedPrice = [Math]::Min($basePrice, $marketAdjusted)
            $suggestedPrice = [Math]::Max($suggestedPrice, $cost * 2)
            $suggestedPrice = [Math]::Round($suggestedPrice / 5) * 5
            
            # Update item
            $updateQuery = @"
            UPDATE Items SET
                CurrentMSRP = @MSRP,
                CurrentMarketAvg = @MarketAvg,
                SuggestedPrice = @SuggestedPrice,
                ModifiedDate = CURRENT_TIMESTAMP
            WHERE ItemID = @ItemID
"@
            $updateParams = @{
                ItemID = $itemID
                MSRP = $msrpPrice
                MarketAvg = $avgPrice
                SuggestedPrice = $suggestedPrice
            }
            
            Invoke-SqliteQuery -DataSource $DatabasePath `
                             -Query $updateQuery `
                             -SqlParameters $updateParams
            
            Write-Host "  Pricing Analysis:" -ForegroundColor Cyan
            Write-Host "    MSRP: `$$msrpPrice" -ForegroundColor White
            Write-Host "    Market Avg: `$$avgPrice" -ForegroundColor White
            Write-Host "    Suggested: `$$suggestedPrice" -ForegroundColor Green
            Write-Host "    Profit Margin: $([Math]::Round(($suggestedPrice/$cost - 1)*100, 1))%" -ForegroundColor Green
        }
        else {
            Write-Host "  No pricing data found" -ForegroundColor Yellow
        }
    }
    
    Write-Host "`n=== Processing Complete ===" -ForegroundColor Green
    
    # Generate summary
    New-SummaryReport -DatabasePath $DatabasePath
}

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
        Write-Host "Total Investment: `$$([Math]::Round($stats.TotalInvestment, 2))" -ForegroundColor Yellow
        Write-Host "Potential Revenue: `$$([Math]::Round($stats.PotentialRevenue, 2))" -ForegroundColor Green
        Write-Host "Average Profit Margin: $([Math]::Round($stats.AvgProfitMargin, 1))%" -ForegroundColor Green
    }
}

#endregion

#region Main Execution

try {
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
        exit 0
    }
    
    # Process data
    Invoke-PurchaseProcessing -CSVPath $InputCSV `
                             -DatabasePath $DatabasePath `
                             -UseAPI:$UseAPI `
                             -UpdateExisting:$UpdateExisting
    
    Write-Host "`nScript completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}

#endregion