<#
.SYNOPSIS
    Extracts lot URLs from auction pages, groups them by auction/catalog, and allows selective extraction.

.DESCRIPTION
    This script processes saved HTML file(s) from auction won/bid pages, identifies different
    auctions by their catalog IDs, and allows you to extract URLs for specific auctions only.
    
    The script will:
    - Process single HTML file OR multiple HTML files in a folder
    - Identify all different auctions across all pages
    - Show lot number ranges for each auction
    - Allow you to select which auction(s) to extract
    - Create separate CSV files for each auction if desired
    - Generate detailed log files of all operations

.PARAMETER Source
    Mandatory. Either:
    - Path to a single HTML file
    - Path to a folder containing multiple HTML files (for paginated results)

.PARAMETER OutputFolder
    Optional. Folder where CSV files will be saved. Default is current directory.

.PARAMETER ExtractMode
    Optional. How to extract auctions:
    - 'Interactive' (default): Shows all auctions and lets you choose
    - 'All': Extracts all auctions to separate CSV files
    - 'Combined': Extracts all auctions to a single CSV file

.PARAMETER AuctionFilter
    Optional. Specific auction catalog ID to extract (e.g., "brolyn10245").
    Use this when you know exactly which auction you want.

.PARAMETER FilePattern
    Optional. Pattern to match HTML files when processing a folder.
    Default is "*.html" to process all HTML files.

.PARAMETER SkipLotsWithoutNumbers
    Optional switch. When set, only extracts lots where lot numbers can be found.
    By default (when not set), includes all lots even without lot numbers.

.PARAMETER ShowSkipped
    Optional switch. When enabled, shows details about lots that were skipped due to missing lot numbers.

.PARAMETER NoLog
    Optional switch. When enabled, disables creation of log file.

.PARAMETER LogPath
    Optional. Custom path for the log file. If not specified, creates log in OutputFolder
    with timestamp (e.g., ExtractAuctions_20241205_143022.log).

.INPUTS
    System.String
    Accepts a string path to either an HTML file or a directory containing HTML files.

.OUTPUTS
    CSV files containing URLs and lot numbers for selected auctions.
    Log file with detailed extraction information (unless -NoLog is specified).

.EXAMPLE
    .\Extract-AuctionURLs.ps1 -Source "won_auctions.html"
    
    Process a single HTML file interactively, including all lots.

.EXAMPLE
    .\Extract-AuctionURLs.ps1 -Source "C:\Auctions\Won" -SkipLotsWithoutNumbers
    
    Process all HTML files in the Won folder, only including lots with lot numbers.

.EXAMPLE
    .\Extract-AuctionURLs.ps1 -Source "C:\Auctions\Won" -FilePattern "LotsWon*.html"
    
    Process only HTML files matching the pattern in the folder.

.EXAMPLE
    .\Extract-AuctionURLs.ps1 -Source "C:\Auctions\Won" -ShowSkipped
    
    Process files and show details about any lots skipped due to missing lot numbers.

.EXAMPLE
    .\Extract-AuctionURLs.ps1 -Source "C:\Auctions\Won" -ExtractMode All -OutputFolder "C:\CSV"
    
    Process all HTML files and extract all auctions to separate CSV files.

.EXAMPLE
    .\Extract-AuctionURLs.ps1 -Source "won.html" -LogPath "C:\Logs\extraction.log"
    
    Process with custom log file location.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 10/03/2025
    Version: 2.3
    Change Date: 10/06/2025
    
    The script identifies auctions by their catalog ID in the URL and groups lots accordingly.
    When processing multiple files (paginated results), it combines all pages before grouping.
    
    VERSION 2.3 CHANGES:
    - Fixed PSScriptAnalyzer warnings
    - Renamed RequireLotNumbers parameter to SkipLotsWithoutNumbers (proper switch behavior)
    - Fixed automatic variable $matches usage
    
    VERSION 2.2 CHANGES:
    - Added comprehensive logging functionality
    - Log file includes all operations, errors, and skipped lots
    - Added -NoLog parameter to disable logging
    - Added -LogPath parameter for custom log location
    
    VERSION 2.1 CHANGES:
    - Now requires lot numbers by default (only extracts lots with identifiable lot numbers)
    - Added parameter to control this behavior
    - Added -ShowSkipped parameter to see details about skipped lots
    - Provides accurate lot counts matching what BidSpotter shows

.LINK
    https://github.com/JONeillSr/Extract-AuctionURLs
#>

param(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Path '$_' does not exist."
        }
        return $true
    })]
    [string]$Source,
    
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_ -PathType Container)) {
            # Try to create the directory if it doesn't exist
            try {
                New-Item -ItemType Directory -Path $_ -Force | Out-Null
                return $true
            }
            catch {
                throw "Cannot create output folder: $_"
            }
        }
        return $true
    })]
    [string]$OutputFolder = ".",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('Interactive', 'All', 'Combined')]
    [string]$ExtractMode = 'Interactive',
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$AuctionFilter = "",
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$FilePattern = "*.html",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipLotsWithoutNumbers,  # FIXED: Renamed from RequireLotNumbers and removed default value
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowSkipped,
    
    [Parameter(Mandatory=$false)]
    [switch]$NoLog,
    
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if ($_) {
            $dir = Split-Path $_ -Parent
            if ($dir -and -not (Test-Path $dir)) {
                try {
                    New-Item -ItemType Directory -Path $dir -Force | Out-Null
                }
                catch {
                    throw "Cannot create log directory: $dir"
                }
            }
        }
        return $true
    })]
    [string]$LogPath = ""
)

# Initialize logging
$script:LogFile = ""
if (-not $NoLog) {
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogFile = Join-Path $OutputFolder "ExtractAuctions_$timestamp.log"
    } else {
        $script:LogFile = $LogPath
    }
    
    # Create log directory if needed
    $logDir = Split-Path $script:LogFile -Parent
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Initialize log file with header
    $logHeader = @"
========================================
Auction URL Extraction Log
========================================
Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script Version: 2.3
PowerShell Version: $($PSVersionTable.PSVersion)
User: $env:USERNAME
Computer: $env:COMPUTERNAME

Parameters:
  Source: $Source
  Output Folder: $OutputFolder
  Extract Mode: $ExtractMode
  Auction Filter: $(if ($AuctionFilter) { $AuctionFilter } else { "None" })
  File Pattern: $FilePattern
  Skip Lots Without Numbers: $SkipLotsWithoutNumbers
  Show Skipped: $ShowSkipped
========================================

"@
    $logHeader | Out-File -FilePath $script:LogFile -Encoding UTF8
}

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug')]
        [string]$Level = 'Info',
        [switch]$NoConsole,
        [switch]$ConsoleOnly
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file if logging is enabled and not console only
    if (-not $ConsoleOnly -and $script:LogFile) {
        $logMessage | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    }
    
    # Write to console unless suppressed
    if (-not $NoConsole) {
        switch ($Level) {
            'Error' { Write-Host $Message -ForegroundColor Red }
            'Warning' { Write-Host $Message -ForegroundColor Yellow }
            'Success' { Write-Host $Message -ForegroundColor Green }
            'Debug' { 
                if ($ShowSkipped) {
                    Write-Host $Message -ForegroundColor Gray 
                }
            }
            default { Write-Host $Message }
        }
    }
}

# Function to extract catalog ID from BidSpotter URL
function Get-CatalogId {
    param([string]$Url)
    
    # FIXED: Using different variable name instead of automatic $matches
    if ($Url -match 'catalogue-id-([^/]+)') {
        $catalogMatch = $Matches[1]
        return $catalogMatch
    }
    return "unknown"
}

# Function to extract lot number from HTML near a URL
function Get-LotNumberNearUrl {
    param(
        [string]$HtmlContent,
        [string]$Url
    )
    
    $urlEscaped = [regex]::Escape($Url)
    
    # FIXED: Store regex matches in different variables to avoid overwriting automatic $matches
    # Look for lot number patterns near the URL
    if ($HtmlContent -match "(?s)$urlEscaped.{0,500}Lot\s*#?\s*(\d+)") {
        $lotMatch = $Matches[1]
        return $lotMatch
    }
    elseif ($HtmlContent -match "(?s)Lot\s*#?\s*(\d+).{0,500}$urlEscaped") {
        $lotMatch = $Matches[1]
        return $lotMatch
    }
    elseif ($HtmlContent -match "(?s)$urlEscaped.{0,500}>(\d+)<") {
        # Sometimes lot number is just in a nearby element
        $lotMatch = $Matches[1]
        return $lotMatch
    }
    
    return ""
}

# Function to extract all lot URLs with their catalog IDs
function Get-AuctionLots {
    param(
        [string]$HtmlContent,
        [bool]$RequireLotNum = $false,
        [bool]$ShowSkippedLots = $false
    )
    
    $auctions = @{}
    $skippedLots = @()
    
    # Find all BidSpotter lot URLs
    $patterns = @(
        'href="(/en-us/auction-catalogues/[^"]+/catalogue-id-[^"]+/lot-[^"]+)"',
        'href="(https?://www\.bidspotter\.com/en-us/auction-catalogues/[^"]+/catalogue-id-[^"]+/lot-[^"]+)"'
    )
    
    $foundUrls = @{}
    
    foreach ($pattern in $patterns) {
        # FIXED: Store regex matches in a local variable
        $urlMatches = [regex]::Matches($HtmlContent, $pattern)
        foreach ($match in $urlMatches) {
            $url = $match.Groups[1].Value
            
            # Ensure full URL
            if ($url -notmatch '^https?://') {
                $url = "https://www.bidspotter.com$url"
            }
            
            # Skip if already processed
            if ($foundUrls.ContainsKey($url)) {
                continue
            }
            $foundUrls[$url] = $true
            
            # Get catalog ID
            $catalogId = Get-CatalogId -Url $url
            
            # Get lot number if possible
            $lotNumber = Get-LotNumberNearUrl -HtmlContent $HtmlContent -Url $url
            
            # Skip if lot number is required but not found
            if ($RequireLotNum -and [string]::IsNullOrWhiteSpace($lotNumber)) {
                $skippedLots += [PSCustomObject]@{
                    URL = $url
                    CatalogId = $catalogId
                    Reason = "No lot number found"
                }
                Write-Log "Skipped lot (no number): $url" -Level Debug
                continue
            }
            
            # Initialize auction group if needed
            if (-not $auctions.ContainsKey($catalogId)) {
                $auctions[$catalogId] = @{
                    CatalogId = $catalogId
                    Lots = @()
                    MinLot = [int]::MaxValue
                    MaxLot = 0
                    AuctionName = ""
                }
            }
            
            # Add lot to auction
            $lotInfo = [PSCustomObject]@{
                URL = $url
                LotNumber = $lotNumber
                CatalogId = $catalogId
            }
            
            $auctions[$catalogId].Lots += $lotInfo
            
            # Track lot number range
            if ($lotNumber -and $lotNumber -match '^\d+$') {
                $lotNum = [int]$lotNumber
                if ($lotNum -lt $auctions[$catalogId].MinLot) {
                    $auctions[$catalogId].MinLot = $lotNum
                }
                if ($lotNum -gt $auctions[$catalogId].MaxLot) {
                    $auctions[$catalogId].MaxLot = $lotNum
                }
            }
        }
    }
    
    # Log skipped lots summary
    if ($skippedLots.Count -gt 0) {
        Write-Log "Skipped $($skippedLots.Count) lot(s) without lot numbers" -Level Warning -NoConsole
        if ($ShowSkippedLots) {
            $skippedByAuction = $skippedLots | Group-Object CatalogId
            foreach ($group in $skippedByAuction) {
                Write-Log "  Auction $($group.Name): $($group.Count) lot(s) skipped" -Level Debug
            }
        }
    }
    
    # Try to extract auction names if possible
    foreach ($catalogId in $auctions.Keys) {
        if ($HtmlContent -match "catalogue-id-$catalogId[^>]+>([^<]+)<") {
            $nameMatch = $Matches[1].Trim()
            $auctions[$catalogId].AuctionName = $nameMatch
        }
    }
    
    # Return both auctions and statistics
    return @{
        Auctions = $auctions
        SkippedCount = $skippedLots.Count
        SkippedLots = $skippedLots
    }
}

# Function to display auction summary
function Show-AuctionSummary {
    param($Auctions)
    
    Write-Host ""
    Write-Host "Found Auctions Summary" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Log "Auction Summary:" -Level Info -NoConsole
    
    $index = 1
    $auctionList = @()
    
    foreach ($catalogId in $Auctions.Keys) {
        $auction = $Auctions[$catalogId]
        $lotCount = $auction.Lots.Count
        
        $lotRange = "Unknown"
        if ($auction.MinLot -ne [int]::MaxValue -and $auction.MaxLot -gt 0) {
            $lotRange = "Lots $($auction.MinLot)-$($auction.MaxLot)"
        }
        
        Write-Host "[$index] Catalog: $catalogId" -ForegroundColor Yellow
        Write-Host "    Lots found: $lotCount" -ForegroundColor Gray
        Write-Host "    Lot range: $lotRange" -ForegroundColor Gray
        
        Write-Log "  [$index] Catalog: $catalogId | Lots: $lotCount | Range: $lotRange" -Level Info -NoConsole
        
        if ($auction.AuctionName) {
            Write-Host "    Name: $($auction.AuctionName)" -ForegroundColor Gray
            Write-Log "    Name: $($auction.AuctionName)" -Level Info -NoConsole
        }
        
        Write-Host ""
        
        $auctionList += @{
            Index = $index
            CatalogId = $catalogId
            Auction = $auction
        }
        
        $index++
    }
    
    return $auctionList
}

# Main script
Write-Host "Auction URL Extractor v2.3" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan
Write-Host ""

if (-not $NoLog) {
    Write-Host "Log file: $script:LogFile" -ForegroundColor Gray
    Write-Host ""
}

Write-Log "Starting extraction process" -Level Info

# Determine if source is a file or folder
$htmlFiles = @()

if (Test-Path -Path $Source -PathType Container) {
    # Source is a folder - get all HTML files
    Write-Log "Processing folder: $Source" -Level Info
    Write-Log "File pattern: $FilePattern" -Level Info
    
    $htmlFiles = Get-ChildItem -Path $Source -Filter $FilePattern -File | Sort-Object Name
    
    if ($htmlFiles.Count -eq 0) {
        Write-Log "ERROR: No HTML files found in folder matching pattern: $FilePattern" -Level Error
        exit 1
    }
    
    Write-Log "Found $($htmlFiles.Count) HTML file(s) to process" -Level Success
    foreach ($file in $htmlFiles) {
        Write-Log "  - $($file.Name)" -Level Info
    }
}
elseif (Test-Path -Path $Source -PathType Leaf) {
    # Source is a single file
    Write-Log "Processing single file: $Source" -Level Info
    $htmlFiles = @(Get-Item -Path $Source)
}
else {
    Write-Log "ERROR: Source not found: $Source" -Level Error
    exit 1
}

# Process all HTML files and combine results
Write-Host ""
Write-Log "Reading HTML file(s)..." -Level Info

$allAuctions = @{}
$fileCount = 0
$totalSkipped = 0
$allSkippedLots = @()

foreach ($file in $htmlFiles) {
    $fileCount++
    Write-Log "[$fileCount/$($htmlFiles.Count)] Processing: $($file.Name)" -Level Info
    
    try {
        $htmlContent = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        
        # Extract auctions from this file
        # FIXED: Pass the actual switch value
        $result = Get-AuctionLots -HtmlContent $htmlContent -RequireLotNum $SkipLotsWithoutNumbers.IsPresent -ShowSkippedLots $ShowSkipped
        $fileAuctions = $result.Auctions
        $totalSkipped += $result.SkippedCount
        $allSkippedLots += $result.SkippedLots
        
        Write-Log "  Found $(($fileAuctions.Values.Lots | Measure-Object).Count) lots in $($fileAuctions.Count) auction(s)" -Level Info -NoConsole
        
        # Merge with existing auctions
        foreach ($catalogId in $fileAuctions.Keys) {
            if ($allAuctions.ContainsKey($catalogId)) {
                # Merge lots from this file into existing auction
                $allAuctions[$catalogId].Lots += $fileAuctions[$catalogId].Lots
                
                # Update lot number ranges
                if ($fileAuctions[$catalogId].MinLot -lt $allAuctions[$catalogId].MinLot) {
                    $allAuctions[$catalogId].MinLot = $fileAuctions[$catalogId].MinLot
                }
                if ($fileAuctions[$catalogId].MaxLot -gt $allAuctions[$catalogId].MaxLot) {
                    $allAuctions[$catalogId].MaxLot = $fileAuctions[$catalogId].MaxLot
                }
                
                # Keep auction name if found
                if ($fileAuctions[$catalogId].AuctionName -and -not $allAuctions[$catalogId].AuctionName) {
                    $allAuctions[$catalogId].AuctionName = $fileAuctions[$catalogId].AuctionName
                }
            }
            else {
                # Add new auction
                $allAuctions[$catalogId] = $fileAuctions[$catalogId]
            }
        }
    }
    catch {
        Write-Log "ERROR: Failed to read HTML file $($file.Name): $_" -Level Error
        continue
    }
}

$auctions = $allAuctions

# Show skipped lots summary if any were skipped
if ($totalSkipped -gt 0) {
    Write-Host ""
    Write-Log "Skipped Lots Summary: $totalSkipped lot(s) without lot numbers" -Level Warning
    
    if ($ShowSkipped) {
        Write-Host "  To include these lots, run without -SkipLotsWithoutNumbers" -ForegroundColor Gray
    } else {
        Write-Host "  To see details, run with -ShowSkipped" -ForegroundColor Gray
    }
}

if ($auctions.Count -eq 0) {
    Write-Log "ERROR: No auction lots found in the HTML file(s)" -Level Error
    if ($totalSkipped -gt 0) {
        Write-Log "Note: $totalSkipped lot(s) were found but skipped due to missing lot numbers" -Level Warning
        Write-Log "Try running without -SkipLotsWithoutNumbers to include them" -Level Warning
    }
    exit 1
}

# Remove duplicates from each auction's lots
Write-Host ""
Write-Log "Removing duplicate URLs..." -Level Info
foreach ($catalogId in $auctions.Keys) {
    $uniqueUrls = @{}
    $uniqueLots = @()
    
    foreach ($lot in $auctions[$catalogId].Lots) {
        if (-not $uniqueUrls.ContainsKey($lot.URL)) {
            $uniqueUrls[$lot.URL] = $true
            $uniqueLots += $lot
        }
    }
    
    $originalCount = $auctions[$catalogId].Lots.Count
    $auctions[$catalogId].Lots = $uniqueLots
    $duplicatesRemoved = $originalCount - $uniqueLots.Count
    
    if ($duplicatesRemoved -gt 0) {
        Write-Log "  Auction $($catalogId): Removed $duplicatesRemoved duplicate(s)" -Level Info
    }
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Green
Write-Host "  Processed $($htmlFiles.Count) HTML file(s)" -ForegroundColor Gray
Write-Host "  Found $($auctions.Count) different auction(s)" -ForegroundColor Gray

$totalLots = 0
foreach ($catalogId in $auctions.Keys) {
    $totalLots += $auctions[$catalogId].Lots.Count
}
Write-Host "  Total unique lots: $totalLots" -ForegroundColor Gray

if ($totalSkipped -gt 0) {
    Write-Host "  Skipped lots: $totalSkipped (no lot numbers)" -ForegroundColor DarkYellow
    Write-Host "  Actual lots with numbers: $totalLots" -ForegroundColor Green
}

Write-Log "Summary: Files=$($htmlFiles.Count), Auctions=$($auctions.Count), Lots=$totalLots, Skipped=$totalSkipped" -Level Success -NoConsole

# Filter by specific auction if requested
if ($AuctionFilter) {
    if ($auctions.ContainsKey($AuctionFilter)) {
        $filteredAuctions = @{
            $AuctionFilter = $auctions[$AuctionFilter]
        }
        $auctions = $filteredAuctions
        Write-Log "Filtered to auction: $AuctionFilter" -Level Info
    }
    else {
        Write-Log "ERROR: Auction with catalog ID '$AuctionFilter' not found" -Level Error
        exit 1
    }
}

# Handle different extract modes
switch ($ExtractMode) {
    'Interactive' {
        # Show summary and let user choose
        $auctionList = Show-AuctionSummary -Auctions $auctions
        
        Write-Host "Which auction(s) do you want to extract?" -ForegroundColor Cyan
        Write-Host "  Enter auction number(s) separated by commas (e.g., 1,3)" -ForegroundColor Gray
        Write-Host "  Enter 'A' for all auctions" -ForegroundColor Gray
        Write-Host "  Enter 'Q' to quit" -ForegroundColor Gray
        Write-Host ""
        
        $choice = Read-Host "Your choice"
        Write-Log "User choice: $choice" -Level Info -NoConsole
        
        if ($choice -eq 'Q') {
            Write-Log "Extraction cancelled by user" -Level Info
            exit 0
        }
        
        $selectedAuctions = @()
        
        if ($choice -eq 'A') {
            $selectedAuctions = $auctionList
            Write-Log "Selected: All auctions" -Level Info -NoConsole
        }
        else {
            $choices = $choice -split ',' | ForEach-Object { $_.Trim() }
            foreach ($c in $choices) {
                if ($c -match '^\d+$') {
                    $num = [int]$c
                    $selected = $auctionList | Where-Object { $_.Index -eq $num }
                    if ($selected) {
                        $selectedAuctions += $selected
                    }
                }
            }
            Write-Log "Selected: $($selectedAuctions.Count) auction(s)" -Level Info -NoConsole
        }
        
        if ($selectedAuctions.Count -eq 0) {
            Write-Log "ERROR: No valid auctions selected" -Level Error
            exit 1
        }
        
        # Extract selected auctions
        foreach ($item in $selectedAuctions) {
            $catalogId = $item.CatalogId
            $auction = $item.Auction
            $outputFile = Join-Path $OutputFolder "loturls_$catalogId.csv"
            
            Write-Host ""
            Write-Log "Extracting auction: $catalogId" -Level Info
            Write-Log "Output file: $outputFile" -Level Info
            
            try {
                $auction.Lots | Select-Object URL, LotNumber | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
                Write-Log "SUCCESS: Saved $($auction.Lots.Count) lots to $outputFile" -Level Success
            }
            catch {
                Write-Log "ERROR: Failed to save CSV: $_" -Level Error
            }
        }
    }
    
    'All' {
        Write-Log "Extract mode: All auctions to separate files" -Level Info -NoConsole
        
        # Extract all auctions to separate files
        foreach ($catalogId in $auctions.Keys) {
            $auction = $auctions[$catalogId]
            $outputFile = Join-Path $OutputFolder "loturls_$catalogId.csv"
            
            Write-Host ""
            Write-Log "Extracting auction: $catalogId" -Level Info
            Write-Log "Output file: $outputFile" -Level Info
            
            try {
                $auction.Lots | Select-Object URL, LotNumber | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
                Write-Log "SUCCESS: Saved $($auction.Lots.Count) lots to $outputFile" -Level Success
            }
            catch {
                Write-Log "ERROR: Failed to save CSV: $_" -Level Error
            }
        }
    }
    
    'Combined' {
        Write-Log "Extract mode: All auctions to single file" -Level Info -NoConsole
        
        # Extract all auctions to a single file
        $outputFile = Join-Path $OutputFolder "all_auctions.csv"
        $allLots = @()
        
        foreach ($catalogId in $auctions.Keys) {
            $allLots += $auctions[$catalogId].Lots
        }
        
        Write-Host ""
        Write-Log "Extracting all auctions to single file" -Level Info
        Write-Log "Output file: $outputFile" -Level Info
        
        try {
            $allLots | Select-Object URL, LotNumber, CatalogId | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
            Write-Log "SUCCESS: Saved $($allLots.Count) total lots from $($auctions.Count) auction(s)" -Level Success
        }
        catch {
            Write-Log "ERROR: Failed to save CSV: $_" -Level Error
        }
    }
}

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host "Extraction complete!" -ForegroundColor Green

Write-Log "Extraction completed successfully" -Level Success -NoConsole
Write-Log "========================================" -Level Info -NoConsole

Write-Host ""
Write-Host "Next step:" -ForegroundColor Yellow
Write-Host "Use Download-ImagesFromURLs.ps1 with the generated CSV file(s)" -ForegroundColor Gray

if (-not $NoLog) {
    Write-Host ""
    Write-Host "Log file saved: $script:LogFile" -ForegroundColor Gray
}
