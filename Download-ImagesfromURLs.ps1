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
    Version: 1.0
    Change Purpose: Initial release
    
    Prerequisites:
        - PowerShell 5.1 or later
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
    https://github.com/[YourUsername]/Download-ImagesFromURLs
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxImagesPerPage = 0,  # 0 means no limit
    
    [Parameter(Mandatory=$false)]
    [switch]$StrictMode,  # Only get images from primary containers
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowDetails  # Show detailed information about found images
)

# Function to extract lot number from the page content
function Get-LotNumber {
    param(
        [string]$HtmlContent,
        [string]$Url
    )
    
    # Try multiple patterns to find lot number
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
    
    # Try to extract from URL as last resort
    if ($Url -match '/lot[_-]?(\d+)' -or $Url -match '[?&]lot=(\d+)') {
        return $matches[1]
    }
    
    Write-Warning "Could not find lot number in page HTML, will use incremental naming"
    return $null
}

# Function to download an image
function Download-Image {
    param(
        [string]$ImageUrl,
        [string]$OutputPath,
        [string]$BaseUrl
    )
    
    try {
        # Handle relative URLs
        if ($ImageUrl -notmatch '^https?://') {
            if ($ImageUrl -match '^//') {
                # Protocol-relative URL
                $ImageUrl = "https:$ImageUrl"
            }
            elseif ($ImageUrl -match '^/') {
                # Absolute path
                $domain = [System.Uri]::new($BaseUrl).GetLeftPart([System.UriPartial]::Authority)
                $ImageUrl = "$domain$ImageUrl"
            }
            else {
                # Relative path
                $basePath = $BaseUrl -replace '[^/]+$', ''
                $ImageUrl = "$basePath$ImageUrl"
            }
        }
        
        # Download the image
        Invoke-WebRequest -Uri $ImageUrl -OutFile $OutputPath -UseBasicParsing
        Write-Host "  Downloaded: $(Split-Path $OutputPath -Leaf)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "  Failed to download image: $ImageUrl"
        Write-Warning "  Error: $_"
        return $false
    }
}

# Function to extract BidSpotter images
function Get-BidSpotterImages {
    param(
        [string]$HtmlContent,
        [bool]$DebugMode
    )
    
    $tempImages = @()
    
    # Strategy 1: Look for the primary lot detail container
    $lotImageContainerPatterns = @(
        '(?s)<div[^>]+class="[^"]*lot-detail[^"]*"[^>]*>(.*?)(?:<div[^>]+class="[^"]*(?:related|similar|other-lots|recommendations)|$)',
        '(?s)<div[^>]+id="lot-images[^"]*"[^>]*>(.*?)</div>\s*</div>',
        '(?s)<section[^>]+class="[^"]*lot-images[^"]*"[^>]*>(.*?)</section>',
        '(?s)<div[^>]+class="[^"]*(?:image-gallery|photo-gallery|item-images)[^"]*"[^>]*>(.*?)(?:<div[^>]+class="[^"]*(?:lot-info|description|related)|</article>|$)'
    )
    
    $lotImagesFound = $false
    foreach ($pattern in $lotImageContainerPatterns) {
        if ($HtmlContent -match $pattern) {
            $lotImageSection = $matches[1]
            if ($DebugMode) {
                Write-Host "    [DEBUG] Found lot image container" -ForegroundColor DarkGray
            }
            
            # Extract all images from this specific section
            $imageExtractionPatterns = @(
                'href="([^"]+\.(?:jpg|jpeg|png|gif|webp)[^"]*)"',
                'data-(?:large|full|zoom|original)[^=]*="([^"]+\.(?:jpg|jpeg|png|gif|webp)[^"]*)"',
                'src="([^"]+\.(?:jpg|jpeg|png|gif|webp)[^"]*)"'
            )
            
            foreach ($imgPattern in $imageExtractionPatterns) {
                $imgMatches = [regex]::Matches($lotImageSection, $imgPattern)
                foreach ($imgMatch in $imgMatches) {
                    $imgUrl = $imgMatch.Groups[1].Value.Replace('&amp;', '&').Replace('\/','/')
                    
                    # Skip obvious thumbnails and UI elements
                    if ($imgUrl -notmatch '(thumb|_tn\.|_sm\.|icon|logo|button|arrow|placeholder|blank|loading)') {
                        # Skip small thumbnail sizes (common thumbnail dimensions)
                        if ($imgUrl -notmatch '[?&][wh]=(?:50|75|100|120|150|155|160|200)\b') {
                            # Check if this looks like a product image
                            if ($imgUrl -match '/(?:lots?|items?|products?|images?|gallery|uploads)/') {
                                $tempImages += $imgUrl
                                if ($DebugMode) {
                                    Write-Host "    [DEBUG] Found lot image: $imgUrl" -ForegroundColor DarkGray
                                }
                            }
                        } elseif ($DebugMode) {
                            Write-Host "    [DEBUG] Skipped thumbnail: $imgUrl" -ForegroundColor DarkGray
                        }
                    }
                }
            }
            
            if ($tempImages.Count -gt 0) {
                $lotImagesFound = $true
                break
            }
        }
    }
    
    # Strategy 2: Look for structured data
    if ($tempImages.Count -eq 0) {
        if ($HtmlContent -match '<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>') {
            $jsonLd = $matches[1]
            if ($jsonLd -match '"@type":\s*"Product"' -or $jsonLd -match '"@type":\s*"Offer"') {
                if ($jsonLd -match '"image":\s*\[([^\]]+)\]') {
                    $imageArray = $matches[1]
                    $jsonImages = [regex]::Matches($imageArray, '"([^"]+)"')
                    foreach ($jsonImg in $jsonImages) {
                        $imgUrl = $jsonImg.Groups[1].Value.Replace('\/','/')
                        # Skip small sizes
                        if ($imgUrl -match '\.(?:jpg|jpeg|png|gif|webp)' -and 
                            $imgUrl -notmatch '[?&][wh]=(?:50|75|100|120|150|155|160|200)\b') {
                            $tempImages += $imgUrl
                            if ($DebugMode) {
                                Write-Host "    [DEBUG] Found JSON-LD image: $imgUrl" -ForegroundColor DarkGray
                            }
                        }
                    }
                }
            }
        }
    }
    
    return $tempImages
}

# Main script
Write-Host "Web Image Downloader" -ForegroundColor Cyan
Write-Host "====================" -ForegroundColor Cyan
Write-Host ""

# Validate CSV file exists
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

# Create output folder if it doesn't exist
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Yellow
}

# Read CSV file
try {
    $urls = Import-Csv $CsvPath
    
    # Check if CSV has URL column, if not assume first column is URL
    if ($urls.Count -gt 0) {
        $urlColumnName = $null
        $firstRow = $urls[0]
        
        # Look for common URL column names
        $possibleColumnNames = @('URL', 'Url', 'url', 'Link', 'link', 'Address', 'WebAddress')
        foreach ($colName in $possibleColumnNames) {
            if ($firstRow.PSObject.Properties.Name -contains $colName) {
                $urlColumnName = $colName
                break
            }
        }
        
        # If no standard column name found, use the first column
        if (-not $urlColumnName) {
            $urlColumnName = ($firstRow.PSObject.Properties.Name)[0]
            Write-Host "Using column '$urlColumnName' for URLs" -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Error "Failed to read CSV file: $_"
    exit 1
}

$totalUrls = $urls.Count
$currentUrl = 0
$fallbackLotNumber = 1000  # Starting number for lots without identifiable numbers

Write-Host "Found $totalUrls URLs to process" -ForegroundColor Cyan
Write-Host ""

foreach ($row in $urls) {
    $currentUrl++
    $url = $row.$urlColumnName
    
    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-Warning "Empty URL at row $currentUrl, skipping..."
        continue
    }
    
    Write-Host "[$currentUrl/$totalUrls] Processing: $url" -ForegroundColor Yellow
    
    try {
        # Fetch the webpage
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing
        $htmlContent = $response.Content
        
        # Extract lot number
        $lotNumber = Get-LotNumber -HtmlContent $htmlContent -Url $url
        
        # If no lot number found, use fallback numbering
        if (-not $lotNumber) {
            $lotNumber = $fallbackLotNumber
            $fallbackLotNumber++
            Write-Host "  Using fallback lot number: $lotNumber" -ForegroundColor Yellow
        } else {
            Write-Host "  Lot Number: $lotNumber" -ForegroundColor Green
        }
        
        # Initialize collections
        $foundImages = @()
        
        # Special handling for BidSpotter
        if ($url -match 'bidspotter\.com') {
            Write-Host "  Detected BidSpotter page - using specialized extraction" -ForegroundColor Cyan
            
            if ($StrictMode -or $ShowDetails) {
                Write-Host "  Strict mode enabled - extracting only current lot images" -ForegroundColor Yellow
            }
            
            $tempImages = Get-BidSpotterImages -HtmlContent $htmlContent -DebugMode $ShowDetails
            
            # Process and deduplicate the found images
            $processedImages = @{}
            foreach ($img in $tempImages) {
                # For Azure CDN images, normalize by removing size parameters
                $baseImageUrl = $img -replace '\?.*$', ''
                
                if (-not $processedImages.ContainsKey($baseImageUrl)) {
                    $processedImages[$baseImageUrl] = @($img)
                } else {
                    $processedImages[$baseImageUrl] += $img
                }
            }
            
            # For each base URL, pick the best quality version
            foreach ($baseUrl in $processedImages.Keys) {
                $versions = $processedImages[$baseUrl]
                $bestVersion = $versions[0]
                
                # If we have multiple versions, pick the largest
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
                    if ($ShowDetails -and $versions.Count -gt 1) {
                        Write-Host "    [DEBUG] Found $($versions.Count) versions, selected largest: $bestVersion" -ForegroundColor DarkGray
                    }
                } else {
                    if ($ShowDetails) {
                        Write-Host "    [DEBUG] Single version found: $bestVersion" -ForegroundColor DarkGray
                    }
                }
                
                $foundImages += $bestVersion
            }
        }
        # For non-BidSpotter sites
        else {
            if (-not $StrictMode) {
                Write-Host "  Using general image extraction" -ForegroundColor Yellow
                
                # Find image URLs in the HTML using general patterns
                $imagePatterns = @(
                    '<img[^>]+src="([^"]+)"[^>]*>',
                    '<img[^>]+src=''([^'']+)''[^>]*>',
                    'data-src="([^"]+)"',
                    'data-lazy="([^"]+)"'
                )
                
                $imageUrlsSet = @{}
                
                foreach ($pattern in $imagePatterns) {
                    $matches = [regex]::Matches($htmlContent, $pattern)
                    foreach ($match in $matches) {
                        $imgUrl = $match.Groups[1].Value.Replace('&amp;', '&').Replace('\/','/')
                        
                        # Filter for likely product images
                        if ($imgUrl -match '\.(jpg|jpeg|png|gif|webp)(\?|$)' -and 
                            $imgUrl -notmatch '(icon|logo|banner|button|arrow|sprite|thumb|_tn\.|_small\.|_xs\.|flag|badge|watermark|avatar|profile|social|share|print|email)') {
                            
                            # Normalize for duplicate detection
                            $normalizedUrl = $imgUrl -replace '\?.*$', ''
                            
                            if (-not $imageUrlsSet.ContainsKey($normalizedUrl)) {
                                $imageUrlsSet[$normalizedUrl] = $imgUrl
                                $foundImages += $imgUrl
                                
                                if ($ShowDetails) {
                                    Write-Host "    [DEBUG] Found image: $imgUrl" -ForegroundColor DarkGray
                                }
                            }
                        }
                    }
                }
            }
        }
        
        if ($ShowDetails) {
            Write-Host "    [DEBUG] Final image count: $($foundImages.Count)" -ForegroundColor DarkGray
        }
        
        Write-Host "  Found $($foundImages.Count) images" -ForegroundColor Cyan
        
        if ($foundImages.Count -eq 0) {
            Write-Warning "  No images found on this page"
            continue
        }
        
        # Apply image limit if specified
        $imagesToDownload = $foundImages
        if ($MaxImagesPerPage -gt 0 -and $foundImages.Count -gt $MaxImagesPerPage) {
            $imagesToDownload = $foundImages | Select-Object -First $MaxImagesPerPage
            Write-Host "  Limiting to first $MaxImagesPerPage images" -ForegroundColor Yellow
        }
        
        # Download images
        $imageCounter = 0
        foreach ($imageUrl in $imagesToDownload) {
            $imageCounter++
            
            # Determine file extension
            $extension = ".jpg"  # Default
            if ($imageUrl -match '\.(\w+)(\?|$)') {
                $extension = ".$($matches[1])"
            }
            
            # Build filename
            if ($imageCounter -eq 1) {
                $filename = "$lotNumber$extension"
            } else {
                $filename = "$lotNumber-$imageCounter$extension"
            }
            
            $outputPath = Join-Path $OutputFolder $filename
            
            # Skip if file already exists
            if (Test-Path $outputPath) {
                Write-Host "  File already exists, skipping: $filename" -ForegroundColor Gray
                continue
            }
            
            # Download the image
            Download-Image -ImageUrl $imageUrl -OutputPath $outputPath -BaseUrl $url
        }
        
        Write-Host "  Completed processing lot $lotNumber" -ForegroundColor Green
        Write-Host ""
        
        # Add a small delay to be respectful to the server
        Start-Sleep -Milliseconds 500
    }
    catch {
        Write-Error "Failed to process URL: $url"
        Write-Error "Error: $_"
        Write-Host ""
    }
}

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Processing complete!" -ForegroundColor Green
Write-Host "Images saved to: $OutputFolder" -ForegroundColor Green

# Summary statistics
$downloadedFiles = Get-ChildItem -Path $OutputFolder -File | Where-Object { $_.Length -gt 0 }
Write-Host "Total files downloaded: $($downloadedFiles.Count)" -ForegroundColor Cyan