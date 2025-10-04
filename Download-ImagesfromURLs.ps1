<#
.SYNOPSIS
    Creates new Outlook profiles for Microsoft 365 from a CSV file containing user UPNs.

.DESCRIPTION
    This script automates the creation of Outlook profiles configured for Microsoft 365 (Exchange Online).
    It reads User Principal Names (UPNs) from a CSV file and creates corresponding Outlook profiles
    with the necessary registry entries for Exchange Online connectivity.
    
    The script creates registry entries in the appropriate Outlook and Windows Messaging Subsystem
    locations to establish the profile structure. Users will still need to complete Modern Authentication
    (OAuth) when first accessing their mailbox.

.PARAMETER CsvPath
    Mandatory. The full path to the CSV file containing user UPNs. The CSV must have a column named 'UPN'
    with email addresses for each user.

.PARAMETER ProfileName
    Optional. The base name for the Outlook profiles. Default is "Microsoft 365 Profile".
    Each profile will be created as "ProfileName - UPN" to ensure uniqueness.

.PARAMETER SetAsDefault
    Optional switch. If specified, the first profile created will be set as the default Outlook profile.
    Only affects the first profile when processing multiple users.

.INPUTS
    CSV file with UPN column containing email addresses.

.OUTPUTS
    Registry entries for Outlook profiles and console output showing creation status.

.EXAMPLE
    .\Create-OutlookProfile.ps1 -CsvPath "C:\Users\Admin\users.csv"
    
    Creates Outlook profiles for all users in the CSV file using the default profile name.

.EXAMPLE
    .\Create-OutlookProfile.ps1 -CsvPath "C:\temp\employees.csv" -ProfileName "Company Email"
    
    Creates profiles with the base name "Company Email" for all users in the CSV file.

.EXAMPLE
    .\Create-OutlookProfile.ps1 -CsvPath "C:\data\users.csv" -SetAsDefault
    
    Creates profiles and sets the first one as the default Outlook profile.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 08/28/2025
    Version: 2.0
    Change Purpose: Converted from COM automation to Microsoft Graph API

    Prerequisites:
                    PowerShell 5.1 or later
                    Administrative privileges recommended
    
    IMPORTANT REQUIREMENTS:
    - Outlook must be closed before running this script
    - Run with Administrative privileges for registry access
    - CSV file must contain 'UPN' column with valid email addresses
    - Targets Office 365/2019/2021 (registry path 16.0)
    - Users must complete Modern Authentication on first Outlook launch
    
    LIMITATIONS:
    - Does not handle complex Exchange configurations
    - Profiles created are basic Exchange Online configurations
    - Does not migrate existing data or settings
    - Registry modifications may require restart of Outlook

.LINK
    https://docs.microsoft.com/en-us/outlook/
    
.LINK
    https://docs.microsoft.com/en-us/microsoft-365/admin/email/
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder
)

# Function to extract lot number from the page content
function Get-LotNumber {
    param(
        [string]$HtmlContent
    )
    
    # Try multiple patterns to find lot number
    $patterns = @(
        'lot-number[^>]*>([^<]+)',
        'lot[^>]*>Lot\s*#?\s*(\d+)',
        'Lot\s*#?\s*(\d+)',
        '"lotNumber"[^:]*:\s*"?(\d+)',
        'data-lot-number="(\d+)"'
    )
    
    foreach ($pattern in $patterns) {
        if ($HtmlContent -match $pattern) {
            return $matches[1].Trim()
        }
    }
    
    # If no lot number found in HTML, try to extract from URL
    # BidSpotter URLs often have the lot ID in them
    Write-Warning "Could not find lot number in page HTML, will use incremental naming"
    return $null
}

# Function to download an image
function Download-Image {
    param(
        [string]$ImageUrl,
        [string]$OutputPath
    )
    
    try {
        # Handle relative URLs
        if ($ImageUrl -notmatch '^https?://') {
            $ImageUrl = "https://www.bidspotter.com$ImageUrl"
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

# Main script
Write-Host "BidSpotter Image Downloader" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan
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
        $lotNumber = Get-LotNumber -HtmlContent $htmlContent
        
        # If no lot number found, use fallback numbering
        if (-not $lotNumber) {
            $lotNumber = $fallbackLotNumber
            $fallbackLotNumber++
            Write-Host "  Using fallback lot number: $lotNumber" -ForegroundColor Yellow
        } else {
            Write-Host "  Lot Number: $lotNumber" -ForegroundColor Green
        }
        
        # Find image URLs in the HTML
        # Multiple patterns to catch different image implementations
        $imagePatterns = @(
            '<img[^>]+src="([^"]+)"[^>]*>',
            '<img[^>]+src=''([^'']+)''[^>]*>',
            'data-src="([^"]+)"',
            'data-lazy="([^"]+)"',
            '"imageUrl":\s*"([^"]+)"',
            '"image":\s*"([^"]+)"',
            '"photos":\s*\[([^\]]+)\]'
        )
        
        $foundImages = @()
        
        foreach ($pattern in $imagePatterns) {
            $matches = [regex]::Matches($htmlContent, $pattern)
            foreach ($match in $matches) {
                $imgUrl = $match.Groups[1].Value
                
                # Filter for likely product images (exclude icons, logos, etc.)
                if ($imgUrl -match '\.(jpg|jpeg|png|gif|webp)' -and 
                    $imgUrl -notmatch '(icon|logo|banner|button|arrow|sprite)') {
                    
                    # Clean up the URL
                    $imgUrl = $imgUrl.Replace('&amp;', '&')
                    
                    # Skip duplicates
                    if ($foundImages -notcontains $imgUrl) {
                        $foundImages += $imgUrl
                    }
                }
            }
        }
        
        # Special handling for JSON arrays of photos
        if ($htmlContent -match '"photos":\s*\[([^\]]+)\]') {
            $photosJson = $matches[1]
            $photoUrls = [regex]::Matches($photosJson, '"([^"]+)"')
            foreach ($photoMatch in $photoUrls) {
                $imgUrl = $photoMatch.Groups[1].Value
                if ($imgUrl -match '\.(jpg|jpeg|png|gif|webp)' -and 
                    $foundImages -notcontains $imgUrl) {
                    $foundImages += $imgUrl
                }
            }
        }
        
        Write-Host "  Found $($foundImages.Count) images" -ForegroundColor Cyan
        
        if ($foundImages.Count -eq 0) {
            Write-Warning "  No images found on this page"
            continue
        }
        
        # Download images
        $imageCounter = 0
        foreach ($imageUrl in $foundImages) {
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
            Download-Image -ImageUrl $imageUrl -OutputPath $outputPath
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