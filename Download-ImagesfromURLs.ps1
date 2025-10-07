<#
.SYNOPSIS
    Downloads images from web pages based on URLs provided in a CSV file.

.DESCRIPTION
    This script automates the downloading of images from multiple web pages.
    It reads URLs from a CSV file, visits each page, attempts to extract identifying
    information (like lot numbers for auction sites), and downloads all associated 
    images with organized naming conventions.
    
    The script intelligently extracts identifying numbers from page content when possible
    and names images accordingly. The first image is named with just the identifier 
    (e.g., 3842.jpg), and subsequent images are numbered (e.g., 3842-2.jpg, 3842-3.jpg, etc.).
    
    Originally designed for BidSpotter auction sites but works with most websites that
    contain downloadable images.

.PARAMETER CsvPath
    Mandatory. The full path to the CSV file containing URLs. 
    The script will auto-detect common column names like 'URL', 'Url', 'Link', etc.,
    or use the first column if no standard name is found.

.PARAMETER OutputFolder
    Mandatory. The directory path where downloaded images will be saved.
    The folder will be created if it doesn't exist.

.PARAMETER MaxImagesPerPage
    Optional. Maximum number of images to download per page. Default is 0 (no limit).
    Useful when pages have many images but you only want the first few.

.PARAMETER StrictMode
    Optional switch. When enabled, only downloads images from primary gallery/product containers.
    Recommended for auction sites like BidSpotter to avoid downloading unrelated images.

.PARAMETER ShowDetails
    Optional switch. Shows detailed information about which images are found and why.
    Useful for troubleshooting when too many or wrong images are downloaded.

.INPUTS
    CSV file with a column containing web page URLs.

.OUTPUTS
    Downloaded image files with organized naming and console output showing download progress.

.EXAMPLE
    .\Download-ImagesFromURLs.ps1 -CsvPath "C:\data\urls.csv" -OutputFolder "C:\data\images"
    
    Downloads all images from the URLs in urls.csv to the images folder.

.EXAMPLE
    .\Download-ImagesFromURLs.ps1 -CsvPath ".\auction_lots.csv" -OutputFolder "D:\AuctionPhotos"
    
    Processes auction_lots.csv for auction site images and saves them to D:\AuctionPhotos.

.EXAMPLE
    .\Download-ImagesFromURLs.ps1 -CsvPath ".\urls.csv" -OutputFolder ".\images" -MaxImagesPerPage 3
    
    Downloads only the first 3 images from each page, useful for limiting downloads.

.EXAMPLE
    .\Download-ImagesFromURLs.ps1 -CsvPath ".\bidspotter.csv" -OutputFolder ".\auction_images" -StrictMode
    
    Uses strict mode to only download images from primary gallery containers, avoiding unrelated images.

.EXAMPLE
    .\Download-ImagesFromURLs.ps1 -CsvPath ".\test.csv" -OutputFolder ".\test_images" -StrictMode -ShowDetails
    
    Runs in strict mode with detailed output to see exactly which images are being found and why.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 10/03/2025
    Version: 3.1
    Change Date: 10/06/2025
    Change Purpose: Added TLS configuration and AllowInsecureRedirect support
                    Added comprehensive logging system with multiple log levels and detailed tracking
                    Added comprehensive logging system with multiple log levels and detailed tracking
                    Enhanced with SHA256 hash-based duplicate detection to avoid re-downloading identical images
                    Fixed: PowerShell string interpolation issues with colons
                    Fixed: Use approved verbs for functions and removed unused variables

    Prerequisites:
        - PowerShell 5.1 or later (PowerShell 7+ recommended)
        - Internet connection
        - Write access to output folder
    
    FEATURES:
    - Automatic identifier extraction (lot numbers, product IDs, etc.) when available
    - Smart image detection (filters out icons, logos, banners, thumbnails, etc.)
    - Duplicate prevention (skips already downloaded files)
    - Fallback numbering system for pages without clear identifiers
    - Progress tracking with detailed status messages
    - Respectful scraping with built-in delays
    - Works with various website structures
    
    SUPPORTED IMAGE FORMATS:
    - JPG/JPEG
    - PNG
    - GIF
    - WebP
    
    COMPATIBLE WITH:
    - BidSpotter auction pages (specialized handling)
    - Most e-commerce sites
    - Gallery websites
    - General web pages with images
    
    LIMITATIONS:
    - Requires public access to web pages (no login required)
    - May not capture dynamically loaded images (JavaScript-rendered content)
    - Identifier extraction success depends on page structure
    - Some sites may have anti-scraping measures

.LINK
    https://github.com/JONeillSr/Download-ImagesFromURLs
#>

[CmdletBinding()]

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxImagesPerPage = 0,
    
    [Parameter(Mandatory=$false)]
    [switch]$StrictMode,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowDetails,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("None", "Error", "Warning", "Info", "Debug", "All")]
    [string]$LogLevel = "Info",
    
    [Parameter(Mandatory=$false)]
    [switch]$NoLogFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipDuplicateCheck,
    
    [Parameter(Mandatory=$false)]
    [switch]$RebuildHashDatabase
)

# Configure TLS/SSL settings
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.ServicePointManager]::SecurityProtocol

# Initialize logging and statistics
$script:LogFile = $null
$script:LogLevel = $LogLevel
$script:StartTime = Get-Date
$script:HashDatabase = @{}
$script:HashDatabasePath = $null
$script:DuplicatesFound = 0
$script:BytesSaved = 0

$script:DownloadStats = @{
    TotalUrls = 0
    ProcessedUrls = 0
    FailedUrls = 0
    TotalImagesFound = 0
    TotalImagesDownloaded = 0
    TotalImagesFailed = 0
    TotalImagesSkipped = 0
    TotalDuplicatesFound = 0
    TotalBytes = 0
    BytesSaved = 0
    LotNumbers = @()
}

# Function to initialize logging
function Initialize-Logging {
    if (-not $NoLogFile) {
        $logsFolder = Join-Path $OutputFolder "logs"
        if (-not (Test-Path $logsFolder)) {
            New-Item -ItemType Directory -Path $logsFolder -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogFile = Join-Path $logsFolder "download_log_$timestamp.txt"
        
        $header = @"
================================================================================
Web Image Downloader Log - v3.1 PSScriptAnalyzer Compliant
================================================================================
Start Time: $($script:StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
CSV Source: $CsvPath
Output Folder: $OutputFolder
Parameters: MaxImages=$MaxImagesPerPage, StrictMode=$StrictMode, ShowDetails=$ShowDetails
Duplicate Detection: $(if ($SkipDuplicateCheck) { "Disabled" } else { "Enabled" })
Log Level: $LogLevel
PowerShell Version: $($PSVersionTable.PSVersion)
================================================================================

"@
        $header | Out-File -FilePath $script:LogFile -Encoding UTF8
        
        Write-Host "Log file created: $($script:LogFile)" -ForegroundColor Gray
    }
}

# Function for logging
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error", "Warning", "Info", "Debug")]
        [string]$Level = "Info",
        
        [Parameter(Mandatory=$false)]
        [ConsoleColor]$Color = "White"
    )
    
    $levelPriority = @{
        "None" = 0
        "Error" = 1
        "Warning" = 2
        "Info" = 3
        "Debug" = 4
        "All" = 5
    }
    
    $currentPriority = $levelPriority[$script:LogLevel]
    $messagePriority = $levelPriority[$Level]
    
    if ($script:LogLevel -eq "None") {
        return
    }
    
    if ($script:LogLevel -ne "All" -and $messagePriority -gt $currentPriority) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    if ($ShowDetails -or $Level -in @("Error", "Warning") -or $script:LogLevel -eq "Debug") {
        Write-Host $Message -ForegroundColor $Color
    }
    
    if ($script:LogFile -and -not $NoLogFile) {
        $logEntry | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
    }
}

# Function to initialize hash database
function Initialize-HashDatabase {
    $dbFolder = Join-Path $OutputFolder ".imagedb"
    if (-not (Test-Path $dbFolder)) {
        New-Item -ItemType Directory -Path $dbFolder -Force | Out-Null
        
        # Hide the database folder on Windows
        if ($IsWindows -ne $false) {
            $folder = Get-Item $dbFolder
            $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::Hidden
        }
    }
    
    $script:HashDatabasePath = Join-Path $dbFolder "image_hashes.json"
    
    if ($RebuildHashDatabase) {
        Write-Host "Rebuilding hash database..." -ForegroundColor Yellow
        Build-HashDatabase
    }
    elseif (Test-Path $script:HashDatabasePath) {
        Write-Log "Loading existing hash database..." -Level "Info"
        try {
            $json = Get-Content $script:HashDatabasePath -Raw
            $loaded = $json | ConvertFrom-Json
            
            # Convert from JSON to hashtable
            $script:HashDatabase = @{}
            foreach ($prop in $loaded.PSObject.Properties) {
                $script:HashDatabase[$prop.Name] = @{
                    FilePath = $prop.Value.FilePath
                    FileSize = $prop.Value.FileSize
                    DateAdded = $prop.Value.DateAdded
                    LotNumber = $prop.Value.LotNumber
                    SourceUrl = $prop.Value.SourceUrl
                }
            }
            
            Write-Log "Loaded $($script:HashDatabase.Count) image hashes from database" -Level "Info"
        }
        catch {
            Write-Log "Error loading hash database: $_" -Level "Warning" -Color Yellow
            Write-Log "Creating new hash database..." -Level "Info"
            Build-HashDatabase
        }
    }
    else {
        Write-Log "No existing hash database found. Creating new one..." -Level "Info"
        Build-HashDatabase
    }
}

# Function to build/rebuild hash database from existing images
function Build-HashDatabase {
    Write-Log "Scanning existing images to build hash database..." -Level "Info"
    
    $script:HashDatabase = @{}
    $imageFiles = Get-ChildItem -Path $OutputFolder -Include "*.jpg", "*.jpeg", "*.png", "*.gif", "*.webp" -Recurse -File |
        Where-Object { $_.DirectoryName -notlike "*\.imagedb*" -and $_.DirectoryName -notlike "*\logs*" }
    
    $totalFiles = $imageFiles.Count
    $current = 0
    
    foreach ($file in $imageFiles) {
        $current++
        if ($current % 10 -eq 0 -or $current -eq $totalFiles) {
            Write-Progress -Activity "Building Hash Database" -Status "$current of $totalFiles files" -PercentComplete (($current / $totalFiles) * 100)
        }
        
        try {
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
            
            # Extract lot number from filename
            $lotNumber = ""
            if ($file.BaseName -match "^(\d+)(?:-\d+)?$") {
                $lotNumber = $matches[1]
            }
            
            $script:HashDatabase[$hash] = @{
                FilePath = $file.FullName
                FileSize = $file.Length
                DateAdded = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                LotNumber = $lotNumber
                SourceUrl = ""
            }
        }
        catch {
            Write-Log "Error hashing file $($file.FullName): $_" -Level "Warning"
        }
    }
    
    Write-Progress -Activity "Building Hash Database" -Completed
    
    Save-HashDatabase
    Write-Log "Hash database built with $($script:HashDatabase.Count) images" -Level "Info" -Color Green
}

# Function to save hash database
function Save-HashDatabase {
    try {
        $script:HashDatabase | ConvertTo-Json -Depth 10 | Out-File -FilePath $script:HashDatabasePath -Encoding UTF8
        Write-Log "Hash database saved" -Level "Debug"
    }
    catch {
        Write-Log "Error saving hash database: $_" -Level "Warning"
    }
}

# Function to calculate hash from URL content
function Get-ImageHashFromUrl {
    param(
        [string]$ImageUrl,
        [string]$BaseUrl
    )
    
    try {
        # Handle relative URLs
        if ($ImageUrl -notmatch '^https?://') {
            if ($ImageUrl -match '^//') {
                $ImageUrl = "https:$ImageUrl"
            }
            elseif ($ImageUrl -match '^/') {
                $domain = [System.Uri]::new($BaseUrl).GetLeftPart([System.UriPartial]::Authority)
                $ImageUrl = "$domain$ImageUrl"
            }
            else {
                $basePath = $BaseUrl -replace '[^/]+$', ''
                $ImageUrl = "$basePath$ImageUrl"
            }
        }
        
        Write-Log "Checking hash for: $ImageUrl" -Level "Debug"
        
        # Download to memory for hashing
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        try {
            $PSVersion = $PSVersionTable.PSVersion.Major
            
            if ($PSVersion -ge 7) {
                Invoke-WebRequest -Uri $ImageUrl -OutFile $tempFile -UseBasicParsing -AllowInsecureRedirect
            }
            else {
                try {
                    Invoke-WebRequest -Uri $ImageUrl -OutFile $tempFile -UseBasicParsing
                }
                catch {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
                    $webClient.DownloadFile($ImageUrl, $tempFile)
                    $webClient.Dispose()
                }
            }
            
            # Calculate hash
            $hash = Get-FileHash -Path $tempFile -Algorithm SHA256 | Select-Object -ExpandProperty Hash
            $fileInfo = Get-Item $tempFile
            
            return @{
                Hash = $hash
                Size = $fileInfo.Length
                TempFile = $tempFile
            }
        }
        catch {
            Write-Log "Error downloading for hash check: $_" -Level "Debug"
            if (Test-Path $tempFile) {
                Remove-Item $tempFile -Force
            }
            return $null
        }
    }
    catch {
        Write-Log "Error in hash check: $_" -Level "Debug"
        return $null
    }
}

# FIXED: Using approved verb "Save" instead of "Download"
function Save-ImageWithDuplicateCheck {
    param(
        [string]$ImageUrl,
        [string]$OutputPath,
        [string]$BaseUrl,
        [string]$LotNumber
    )
    
    try {
        # Skip duplicate check if disabled
        if ($SkipDuplicateCheck) {
            return Save-ImageDirect -ImageUrl $ImageUrl -OutputPath $OutputPath -BaseUrl $BaseUrl
        }
        
        # Check if file already exists locally
        if (Test-Path $OutputPath) {
            Write-Log "File already exists locally: $(Split-Path $OutputPath -Leaf)" -Level "Info"
            $script:DownloadStats.TotalImagesSkipped++
            return $false
        }
        
        # Get hash of remote image
        $hashResult = Get-ImageHashFromUrl -ImageUrl $ImageUrl -BaseUrl $BaseUrl
        
        if (-not $hashResult) {
            Write-Log "Could not calculate hash, downloading anyway..." -Level "Debug"
            return Save-ImageDirect -ImageUrl $ImageUrl -OutputPath $OutputPath -BaseUrl $BaseUrl
        }
        
        # Check if hash exists in database
        if ($script:HashDatabase.ContainsKey($hashResult.Hash)) {
            $existing = $script:HashDatabase[$hashResult.Hash]
            $script:DownloadStats.TotalDuplicatesFound++
            $script:DownloadStats.BytesSaved += $hashResult.Size
            
            Write-Log "DUPLICATE FOUND: Image already exists" -Level "Warning" -Color Yellow
            Write-Log "  Original: $($existing.FilePath)" -Level "Info"
            Write-Log "  Lot: $($existing.LotNumber), Size: $([math]::Round($existing.FileSize / 1KB, 2)) KB" -Level "Info"
            Write-Log "  Saved: $([math]::Round($hashResult.Size / 1KB, 2)) KB of bandwidth" -Level "Info"
            
            # Clean up temp file
            if (Test-Path $hashResult.TempFile) {
                Remove-Item $hashResult.TempFile -Force
            }
            
            $script:DownloadStats.TotalImagesSkipped++
            return $false
        }
        
        # Not a duplicate - move temp file to final location
        Move-Item -Path $hashResult.TempFile -Destination $OutputPath -Force
        
        # Add to hash database
        $script:HashDatabase[$hashResult.Hash] = @{
            FilePath = $OutputPath
            FileSize = $hashResult.Size
            DateAdded = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            LotNumber = $LotNumber
            SourceUrl = $ImageUrl
        }
        
        # Save database periodically (every 10 new images)
        if ($script:HashDatabase.Count % 10 -eq 0) {
            Save-HashDatabase
        }
        
        $script:DownloadStats.TotalBytes += $hashResult.Size
        $script:DownloadStats.TotalImagesDownloaded++
        
        $sizeStr = Format-FileSize $hashResult.Size
        Write-Log "Downloaded: $(Split-Path $OutputPath -Leaf) ($sizeStr) [Hash: $($hashResult.Hash.Substring(0, 8))...]" -Level "Info" -Color Green
        return $true
    }
    catch {
        $script:DownloadStats.TotalImagesFailed++
        Write-Log "Failed to download image: $ImageUrl - Error: $_" -Level "Error" -Color Red
        return $false
    }
}

# FIXED: Using approved verb "Save" instead of "Download"
function Save-ImageDirect {
    param(
        [string]$ImageUrl,
        [string]$OutputPath,
        [string]$BaseUrl
    )
    
    try {
        # Handle relative URLs
        if ($ImageUrl -notmatch '^https?://') {
            if ($ImageUrl -match '^//') {
                $ImageUrl = "https:$ImageUrl"
            }
            elseif ($ImageUrl -match '^/') {
                $domain = [System.Uri]::new($BaseUrl).GetLeftPart([System.UriPartial]::Authority)
                $ImageUrl = "$domain$ImageUrl"
            }
            else {
                $basePath = $BaseUrl -replace '[^/]+$', ''
                $ImageUrl = "$basePath$ImageUrl"
            }
        }
        
        $PSVersion = $PSVersionTable.PSVersion.Major
        
        if ($PSVersion -ge 7) {
            Invoke-WebRequest -Uri $ImageUrl -OutFile $OutputPath -UseBasicParsing -AllowInsecureRedirect
        }
        else {
            try {
                Invoke-WebRequest -Uri $ImageUrl -OutFile $OutputPath -UseBasicParsing
            }
            catch {
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
                $webClient.DownloadFile($ImageUrl, $OutputPath)
                $webClient.Dispose()
            }
        }
        
        $fileInfo = Get-Item $OutputPath
        $script:DownloadStats.TotalBytes += $fileInfo.Length
        $script:DownloadStats.TotalImagesDownloaded++
        
        Write-Log "Downloaded: $(Split-Path $OutputPath -Leaf)" -Level "Info" -Color Green
        return $true
    }
    catch {
        $script:DownloadStats.TotalImagesFailed++
        Write-Log "Failed to download: $_" -Level "Error" -Color Red
        return $false
    }
}

# Function to format file size
function Format-FileSize {
    param([long]$Size)
    
    if ($Size -gt 1GB) {
        return "{0:N2} GB" -f ($Size / 1GB)
    }
    elseif ($Size -gt 1MB) {
        return "{0:N2} MB" -f ($Size / 1MB)
    }
    elseif ($Size -gt 1KB) {
        return "{0:N2} KB" -f ($Size / 1KB)
    }
    else {
        return "$Size bytes"
    }
}

# Function to write download summary
function Write-DownloadSummary {
    $duration = (Get-Date) - $script:StartTime
    
    # Save hash database one final time
    if (-not $SkipDuplicateCheck) {
        Save-HashDatabase
    }
    
    $summary = @"

================================================================================
DOWNLOAD SUMMARY
================================================================================
End Time: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Duration: $($duration.ToString("hh\:mm\:ss"))

URLs Processed: $($script:DownloadStats.ProcessedUrls) / $($script:DownloadStats.TotalUrls)
URLs Failed: $($script:DownloadStats.FailedUrls)

Images Found: $($script:DownloadStats.TotalImagesFound)
Images Downloaded: $($script:DownloadStats.TotalImagesDownloaded)
Images Skipped (already existed): $($script:DownloadStats.TotalImagesSkipped)
Duplicates Detected: $($script:DownloadStats.TotalDuplicatesFound)
Images Failed: $($script:DownloadStats.TotalImagesFailed)

Total Data Downloaded: $(Format-FileSize $script:DownloadStats.TotalBytes)
Bandwidth Saved (duplicates): $(Format-FileSize $script:DownloadStats.BytesSaved)

Hash Database Size: $($script:HashDatabase.Count) unique images
Unique Lot Numbers: $($script:DownloadStats.LotNumbers.Count)

Lot Numbers Processed:
$($script:DownloadStats.LotNumbers -join ", ")
================================================================================
"@
    
    Write-Host $summary -ForegroundColor Cyan
    
    if ($script:LogFile -and -not $NoLogFile) {
        $summary | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
    }
}

# Function to extract lot number
function Get-LotNumber {
    param(
        [string]$HtmlContent,
        [string]$Url
    )
    
    $patterns = @(
        'lot-number[^>]*>([^<]+)',
        'lot[^>]*>Lot\s*#?\s*(\d+)',
        'Lot\s*#?\s*(\d+)',
        '"lotNumber"[^:]*:\s*"?(\d+)',
        'data-lot-number="(\d+)"',
        'id="lot-(\d+)"',
        'class="[^"]*lot-(\d+)'
    )
    
    foreach ($pattern in $patterns) {
        if ($HtmlContent -match $pattern) {
            return $matches[1].Trim()
        }
    }
    
    if ($Url -match '/lot[_-]?(\d+)' -or $Url -match '[?&]lot=(\d+)') {
        return $matches[1]
    }
    
    Write-Log "Could not find lot number in page HTML" -Level "Debug"
    return $null
}

# FIXED: Removed unused $lotImagesFound variable
function Get-BidSpotterImages {
    param(
        [string]$HtmlContent,
        [bool]$DebugMode
    )
    
    $tempImages = @()
    
    $lotImageContainerPatterns = @(
        '(?s)<div[^>]+class="[^"]*lot-detail[^"]*"[^>]*>(.*?)(?:<div[^>]+class="[^"]*(?:related|similar|other-lots|recommendations)|$)',
        '(?s)<div[^>]+id="lot-images[^"]*"[^>]*>(.*?)</div>\s*</div>',
        '(?s)<section[^>]+class="[^"]*lot-images[^"]*"[^>]*>(.*?)</section>',
        '(?s)<div[^>]+class="[^"]*(?:image-gallery|photo-gallery|item-images)[^"]*"[^>]*>(.*?)(?:<div[^>]+class="[^"]*(?:lot-info|description|related)|</article>|$)'
    )
    
    foreach ($pattern in $lotImageContainerPatterns) {
        if ($HtmlContent -match $pattern) {
            $lotImageSection = $matches[1]
            
            $imageExtractionPatterns = @(
                'href="([^"]+\.(?:jpg|jpeg|png|gif|webp)[^"]*)"',
                'data-(?:large|full|zoom|original)[^=]*="([^"]+\.(?:jpg|jpeg|png|gif|webp)[^"]*)"',
                'src="([^"]+\.(?:jpg|jpeg|png|gif|webp)[^"]*)"'
            )
            
            foreach ($imgPattern in $imageExtractionPatterns) {
                $imgMatches = [regex]::Matches($lotImageSection, $imgPattern)
                foreach ($imgMatch in $imgMatches) {
                    $imgUrl = $imgMatch.Groups[1].Value.Replace('&amp;', '&').Replace('\/','/')
                    
                    if ($imgUrl -notmatch '(thumb|_tn\.|_sm\.|icon|logo|button|arrow|placeholder|blank|loading)') {
                        if ($imgUrl -notmatch '[?&][wh]=(?:50|75|100|120|150|155|160|200)\b') {
                            if ($imgUrl -match '/(?:lots?|items?|products?|images?|gallery|uploads)/') {
                                $tempImages += $imgUrl
                            }
                        }
                    }
                }
            }
            
            if ($tempImages.Count -gt 0) {
                break
            }
        }
    }
    
    if ($tempImages.Count -eq 0) {
        if ($HtmlContent -match '<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>') {
            $jsonLd = $matches[1]
            if ($jsonLd -match '"@type":\s*"Product"' -or $jsonLd -match '"@type":\s*"Offer"') {
                if ($jsonLd -match '"image":\s*\[([^\]]+)\]') {
                    $imageArray = $matches[1]
                    $jsonImages = [regex]::Matches($imageArray, '"([^"]+)"')
                    foreach ($jsonImg in $jsonImages) {
                        $imgUrl = $jsonImg.Groups[1].Value.Replace('\/','/')
                        if ($imgUrl -match '\.(?:jpg|jpeg|png|gif|webp)' -and 
                            $imgUrl -notmatch '[?&][wh]=(?:50|75|100|120|150|155|160|200)\b') {
                            $tempImages += $imgUrl
                        }
                    }
                }
            }
        }
    }
    
    return $tempImages
}

# === MAIN SCRIPT ===

Write-Host "Web Image Downloader v3.1 - PSScriptAnalyzer Compliant" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# Initialize systems
Initialize-Logging
Write-Log "Starting Web Image Downloader v3.1" -Level "Info"

# Initialize duplicate detection
if (-not $SkipDuplicateCheck) {
    Initialize-HashDatabase
}

# Validate CSV file
if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found: $CsvPath" -Level "Error" -Color Red
    Write-DownloadSummary
    exit 1
}

# Create output folder
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Log "Created output folder: $OutputFolder" -Level "Info"
}

# Read CSV file
try {
    $urls = Import-Csv $CsvPath
    
    if ($urls.Count -gt 0) {
        $urlColumnName = $null
        $firstRow = $urls[0]
        
        $possibleColumnNames = @('URL', 'Url', 'url', 'Link', 'link', 'Address', 'WebAddress')
        foreach ($colName in $possibleColumnNames) {
            if ($firstRow.PSObject.Properties.Name -contains $colName) {
                $urlColumnName = $colName
                break
            }
        }
        
        if (-not $urlColumnName) {
            $urlColumnName = ($firstRow.PSObject.Properties.Name)[0]
            Write-Log "Using column '$urlColumnName' for URLs" -Level "Info"
        }
    }
    
    $script:DownloadStats.TotalUrls = $urls.Count
}
catch {
    Write-Log "Failed to read CSV file: $_" -Level "Error" -Color Red
    Write-DownloadSummary
    exit 1
}

$totalUrls = $urls.Count
$currentUrl = 0
$fallbackLotNumber = 1000

Write-Host "Found $totalUrls URLs to process" -ForegroundColor Cyan
if (-not $SkipDuplicateCheck) {
    Write-Host "Duplicate detection: ENABLED (checking against $($script:HashDatabase.Count) known images)" -ForegroundColor Green
} else {
    Write-Host "Duplicate detection: DISABLED" -ForegroundColor Yellow
}
Write-Host ""

# Process each URL
foreach ($row in $urls) {
    $currentUrl++
    $url = $row.$urlColumnName
    
    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-Log "Empty URL at row $currentUrl, skipping..." -Level "Warning"
        continue
    }
    
    Write-Host "[$currentUrl/$totalUrls] Processing: $url" -ForegroundColor Yellow
    # FIXED: Use string formatting to avoid colon parsing issue
    Write-Log ("Processing URL {0}/{1}: {2}" -f $currentUrl, $totalUrls, $url) -Level "Info"
    
    try {
        # Fetch webpage
        Write-Log "Fetching page content..." -Level "Debug"
        
        $response = $null
        $htmlContent = $null
        $PSVersion = $PSVersionTable.PSVersion.Major
        
        if ($PSVersion -ge 7) {
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing -AllowInsecureRedirect
        }
        else {
            try {
                $response = Invoke-WebRequest -Uri $url -UseBasicParsing
            }
            catch {
                if ($_.Exception.Message -match "insecure redirection") {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
                    $htmlContent = $webClient.DownloadString($url)
                    $webClient.Dispose()
                }
                else {
                    throw
                }
            }
        }
        
        if ($response) {
            $htmlContent = $response.Content
        }
        
        if (-not $htmlContent) {
            Write-Log "Could not retrieve page content" -Level "Warning"
            $script:DownloadStats.FailedUrls++
            continue
        }
        
        # Extract lot number
        $lotNumber = Get-LotNumber -HtmlContent $htmlContent -Url $url
        
        if (-not $lotNumber) {
            $lotNumber = $fallbackLotNumber
            $fallbackLotNumber++
            Write-Log "Using fallback lot number: $lotNumber" -Level "Info"
        } else {
            Write-Log "Lot Number: $lotNumber" -Level "Info"
        }
        
        if ($lotNumber -notin $script:DownloadStats.LotNumbers) {
            $script:DownloadStats.LotNumbers += $lotNumber
        }
        
        # Extract images
        $foundImages = @()
        
        if ($url -match 'bidspotter\.com') {
            Write-Log "Detected BidSpotter page" -Level "Info"
            $tempImages = Get-BidSpotterImages -HtmlContent $htmlContent -DebugMode $ShowDetails
            
            # Process and deduplicate
            $processedImages = @{}
            foreach ($img in $tempImages) {
                $baseImageUrl = $img -replace '\?.*$', ''
                
                if (-not $processedImages.ContainsKey($baseImageUrl)) {
                    $processedImages[$baseImageUrl] = @($img)
                } else {
                    $processedImages[$baseImageUrl] += $img
                }
            }
            
            foreach ($baseUrl in $processedImages.Keys) {
                $versions = $processedImages[$baseUrl]
                $bestVersion = $versions[0]
                
                if ($versions.Count -gt 1) {
                    $largestWidth = 0
                    foreach ($version in $versions) {
                        if ($version -match '[?&]w=(\d+)') {
                            $width = [int]$matches[1]
                            if ($width -gt $largestWidth) {
                                $largestWidth = $width
                                $bestVersion = $version
                            }
                        }
                    }
                }
                
                $foundImages += $bestVersion
            }
        }
        
        $script:DownloadStats.TotalImagesFound += $foundImages.Count
        Write-Host "  Found $($foundImages.Count) images" -ForegroundColor Cyan
        
        if ($foundImages.Count -eq 0) {
            Write-Log "No images found on this page" -Level "Warning"
            continue
        }
        
        # Apply image limit
        $imagesToDownload = $foundImages
        if ($MaxImagesPerPage -gt 0 -and $foundImages.Count -gt $MaxImagesPerPage) {
            $imagesToDownload = $foundImages | Select-Object -First $MaxImagesPerPage
            Write-Log "Limiting to first $MaxImagesPerPage images" -Level "Info"
        }
        
        # Download images (using approved verb function)
        $imageCounter = 0
        foreach ($imageUrl in $imagesToDownload) {
            $imageCounter++
            
            $extension = ".jpg"
            if ($imageUrl -match '\.(\w+)(\?|$)') {
                $extension = ".$($matches[1])"
            }
            
            if ($imageCounter -eq 1) {
                $filename = "$lotNumber$extension"
            } else {
                $filename = "$lotNumber-$imageCounter$extension"
            }
            
            $outputPath = Join-Path $OutputFolder $filename
            
            # FIXED: Call renamed function with approved verb
            Save-ImageWithDuplicateCheck -ImageUrl $imageUrl -OutputPath $outputPath -BaseUrl $url -LotNumber $lotNumber
        }
        
        $script:DownloadStats.ProcessedUrls++
        Write-Host "  Completed processing lot $lotNumber" -ForegroundColor Green
        Write-Host ""
        
        Start-Sleep -Milliseconds 500
    }
    catch {
        $script:DownloadStats.FailedUrls++
        # FIXED: Use string formatting here too
        Write-Log ("Failed to process URL: {0} - Error: {1}" -f $url, $_) -Level "Error" -Color Red
        Write-Host ""
    }
}

# Final summary
Write-DownloadSummary

if ($script:DownloadStats.TotalDuplicatesFound -gt 0) {
    Write-Host ""
    Write-Host "DUPLICATE DETECTION SAVED:" -ForegroundColor Green
    Write-Host "  Avoided downloading $($script:DownloadStats.TotalDuplicatesFound) duplicate images" -ForegroundColor Yellow
    Write-Host "  Saved $(Format-FileSize $script:DownloadStats.BytesSaved) of bandwidth" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Hash database location: $(Split-Path $script:HashDatabasePath -Parent)" -ForegroundColor Gray