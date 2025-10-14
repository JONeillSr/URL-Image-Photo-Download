<#
.SYNOPSIS
    Shared functions for Purchase Tracking system.

.DESCRIPTION
    Common functions used by both CLI and GUI versions of Purchase Tracking.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10/13/2025
    Version: 1.0.0
    Change Date:
    Change Purpose:

.CHANGELOG
    1.0.0 - Initial module creation with shared functions
#>

#region Helper Functions

function ConvertTo-Hashtable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $InputObject
    )

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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Item
    )

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

#endregion

#region Database Functions

function Initialize-PurchaseDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DatabasePath
    )

    Write-Verbose "Initializing database at: $DatabasePath"

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
        Invoke-SqliteQuery -DataSource $DatabasePath -Query $schema
        Write-Verbose "Database initialized successfully"
    }
    catch {
        Write-Error "Failed to initialize database: $_"
        throw
    }
}

function Get-PurchaseItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DatabasePath,

        [Parameter(Mandatory=$false)]
        [string]$Filter = ""
    )

    $query = @"
        SELECT
            ItemID,
            Description,
            Category,
            Quantity,
            ROUND(PerItemCost, 2) as PerItemCost,
            ROUND(CurrentMarketAvg, 2) as MarketAvg,
            ROUND(SuggestedPrice, 2) as OurPrice,
            ROUND((SuggestedPrice - PerItemCost), 2) as Profit,
            CASE WHEN PerItemCost > 0
                THEN ROUND((SuggestedPrice/PerItemCost - 1) * 100, 1)
                ELSE 0 END as Margin
        FROM Items
        WHERE IsActive = 1
        $(if ($Filter) { "AND Description LIKE '%$Filter%'" })
        ORDER BY ItemID
"@

    return Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
}

function Update-PurchaseItemField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DatabasePath,

        [Parameter(Mandatory=$true)]
        [int]$ItemID,

        [Parameter(Mandatory=$true)]
        [string]$FieldName,

        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Value
    )

    try {
        $query = "UPDATE Items SET $FieldName = @Value, ModifiedDate = CURRENT_TIMESTAMP WHERE ItemID = @ID"
        $params = @{ Value = $Value; ID = $ItemID }
        Invoke-SqliteQuery -DataSource $DatabasePath -Query $query -SqlParameters $params
        return $true
    }
    catch {
        Write-Error "Failed to update field $FieldName for item $ItemID: $_"
        return $false
    }
}

#endregion

#region Web Scraping Functions

function Get-WebContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$URL,

        [Parameter(Mandatory=$false)]
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ProductDescription,

        [Parameter(Mandatory=$true)]
        [string]$Site
    )

    Write-Verbose "Searching $Site for: $ProductDescription"

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
        Write-Warning "Unknown site: $Site"
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
            Write-Verbose "Found average price: $avgPrice"
            return $avgPrice
        }
    }

    Write-Verbose "No prices found on $Site"
    return $null
}

function Invoke-PriceResearch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DatabasePath,

        [Parameter(Mandatory=$true)]
        [int]$ItemID,

        [Parameter(Mandatory=$true)]
        [string]$Description,

        [Parameter(Mandatory=$false)]
        [string[]]$Sites = @('walmart.com', 'amazon.com', 'ebay.com', 'lippert.com'),

        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressCallback
    )

    $foundPrices = @()

    foreach ($site in $Sites) {
        if ($ProgressCallback) {
            & $ProgressCallback "  Searching $site..."
        }

        $price = Search-ProductPrice -ProductDescription $Description -Site $site

        if ($price) {
            $foundPrices += $price

            if ($ProgressCallback) {
                & $ProgressCallback "    Found: `$$price"
            }

            # Store in price history
            $historyQuery = @"
            INSERT INTO PriceHistory (ItemID, Source, Price, Confidence)
            VALUES (@ItemID, @Source, @Price, @Confidence)
"@
            $historyParams = @{
                ItemID = $ItemID
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

    if ($foundPrices.Count -gt 0) {
        $avgPrice = [Math]::Round(($foundPrices | Measure-Object -Average).Average, 2)
        $msrpPrice = [Math]::Round(($foundPrices | Measure-Object -Maximum).Maximum, 2)

        return @{
            Success = $true
            AveragePrice = $avgPrice
            MSRPPrice = $msrpPrice
            PriceCount = $foundPrices.Count
        }
    }

    return @{
        Success = $false
        AveragePrice = 0
        MSRPPrice = 0
        PriceCount = 0
    }
}

#endregion

#region Processing Functions

function Invoke-ItemProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DatabasePath,

        [Parameter(Mandatory=$false)]
        [string]$CSVPath,

        [Parameter(Mandatory=$false)]
        [switch]$UpdateExisting,

        [Parameter(Mandatory=$false)]
        [switch]$PerformPriceLookup,

        [Parameter(Mandatory=$false)]
        [string[]]$PriceLookupSites = @('walmart.com', 'amazon.com', 'ebay.com', 'lippert.com'),

        [Parameter(Mandatory=$false)]
        [scriptblock]$ProgressCallback
    )

    # Load items
    $items = @()
    if ($CSVPath -and (Test-Path $CSVPath)) {
        if ($ProgressCallback) {
            & $ProgressCallback "Loading items from CSV..."
        }
        $items = Import-Csv -Path $CSVPath
    }
    elseif ($UpdateExisting) {
        if ($ProgressCallback) {
            & $ProgressCallback "Loading items from database..."
        }
        $query = "SELECT * FROM Items WHERE IsActive = 1"
        $items = Invoke-SqliteQuery -DataSource $DatabasePath -Query $query
    }

    if ($items.Count -eq 0) {
        if ($ProgressCallback) {
            & $ProgressCallback "No items to process"
        }
        return
    }

    $totalItems = $items.Count
    $currentItem = 0

    foreach ($item in $items) {
        $currentItem++

        if ($ProgressCallback) {
            & $ProgressCallback "[$currentItem/$totalItems] Processing: $($item.Description)"
        }

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
                if ($ProgressCallback) {
                    & $ProgressCallback "  Item already exists (ID: $itemID)"
                }
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

                if ($ProgressCallback) {
                    & $ProgressCallback "  Added to database (ID: $itemID)"
                }
            }
        }
        else {
            $itemID = $normalizedItem.ItemID
        }

        # Search for prices if enabled
        if ($PerformPriceLookup) {
            $priceResult = Invoke-PriceResearch `
                -DatabasePath $DatabasePath `
                -ItemID $itemID `
                -Description $normalizedItem.Description `
                -Sites $PriceLookupSites `
                -ProgressCallback $ProgressCallback

            if ($priceResult.Success) {
                # Calculate recommendations
                $cost = if ($normalizedItem.PerItemCost) {
                    [decimal]$normalizedItem.PerItemCost
                } else { 1.0 }

                $basePrice = $cost * 3.5
                $marketAdjusted = $priceResult.AveragePrice * 0.85
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
                    MSRP = $priceResult.MSRPPrice
                    MarketAvg = $priceResult.AveragePrice
                    SuggestedPrice = $suggestedPrice
                }

                Invoke-SqliteQuery -DataSource $DatabasePath `
                                 -Query $updateQuery `
                                 -SqlParameters $updateParams

                if ($ProgressCallback) {
                    $margin = [Math]::Round(($suggestedPrice/$cost - 1)*100, 1)
                    & $ProgressCallback "  MSRP: `$$($priceResult.MSRPPrice) | Market: `$$($priceResult.AveragePrice) | Suggested: `$$suggestedPrice | Margin: $margin%"
                }
            }
        }
    }

    if ($ProgressCallback) {
        & $ProgressCallback "Processing complete!"
    }
}

#endregion

# Export module members
Export-ModuleMember -Function @(
    'ConvertTo-Hashtable',
    'Normalize-ItemData',
    'Initialize-PurchaseDatabase',
    'Get-PurchaseItem',
    'Update-PurchaseItemField',
    'Get-WebContent',
    'Search-ProductPrice',
    'Invoke-PriceResearch',
    'Invoke-ItemProcessing'
)
