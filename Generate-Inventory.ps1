<#
.SYNOPSIS
    Creates a master inventory of all downloaded auction images with detailed lot information.

.DESCRIPTION
    This script scans a specified root folder containing auction images and generates
    a comprehensive CSV inventory. The inventory includes lot numbers, auction houses,
    file sizes, dates, and image counts for each lot.
    
    The script:
    - Recursively scans for all image files (jpg, jpeg, png, gif, webp)
    - Extracts lot numbers from filenames
    - Groups images by lot (handling multi-image lots like 1234-2.jpg, 1234-3.jpg)
    - Identifies auction house and name from folder structure
    - Calculates storage usage and provides statistics
    - Exports results to a timestamped CSV file
    
    Folder structure expected:
    RootFolder\
    ├── AuctionHouse1\
    │   ├── AuctionName1\
    │   │   ├── 1001.jpg
    │   │   ├── 1001-2.jpg
    │   │   └── 1002.jpg
    │   └── AuctionName2\
    └── AuctionHouse2\

.PARAMETER RootFolder
    Mandatory. The root folder path containing auction images.
    Example: "C:\Auctions"

.PARAMETER OutputFile
    Optional. The name of the output CSV file.
    Default: "auction_inventory_YYYYMMDD.csv" with current date.
    The file will be created in the RootFolder.

.INPUTS
    System.String
    Accepts a string path to the root auction folder.

.OUTPUTS
    CSV file containing the auction inventory with the following columns:
    - LotNumber: Extracted lot number
    - AuctionHouse: First level folder name
    - AuctionName: Second level folder name
    - ImageCount: Total images for this lot
    - FilePath: Full path to the main image
    - FileSize: Size in KB
    - DateDownloaded: File creation date
    - MainImage: Filename of the primary image

.EXAMPLE
    .\Generate-AuctionInventory.ps1 -RootFolder "C:\Auctions"
    
    Generates inventory for all images in C:\Auctions with default output filename.

.EXAMPLE
    .\Generate-AuctionInventory.ps1 -RootFolder "D:\MyAuctions" -OutputFile "inventory_2024.csv"
    
    Creates a custom-named inventory file in D:\MyAuctions.

.EXAMPLE
    .\Generate-AuctionInventory.ps1 -RootFolder "C:\Auctions\2024"
    
    Generates inventory for a specific year's auctions.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 10/06/2025
    Version: 1.1
    
    The script identifies "main" images (without -2, -3 suffixes) to avoid duplicates
    in the inventory. Each lot appears once with the total count of associated images.
    
    VERSION 1.1 CHANGES:
    - Fixed PSScriptAnalyzer warnings (Sort-Object syntax)
    - Added comprehensive PowerShell help documentation
    - Fixed automatic variable $matches usage
    - Improved error handling
    
    File Naming Convention Expected:
    - Main image: LOTNUMBER.ext (e.g., 1234.jpg)
    - Additional images: LOTNUMBER-N.ext (e.g., 1234-2.jpg, 1234-3.jpg)

.LINK
    https://github.com/JONeillSr/Generate-AuctionInventory
#>

param(
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Container)) {
            throw "Root folder path '$_' does not exist."
        }
        return $true
    })]
    [string]$RootFolder,
    
    [Parameter(Mandatory=$false, Position=1)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFile = "auction_inventory_$(Get-Date -Format 'yyyyMMdd').csv"
)

# Initialize
Write-Host "Auction Inventory Generator v1.1" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Root Folder: $RootFolder" -ForegroundColor Gray
Write-Host "Output File: $OutputFile" -ForegroundColor Gray
Write-Host ""

$inventory = @()
$errorCount = 0

# Scan all image files
Write-Host "Scanning for image files..." -ForegroundColor Yellow
$imageFiles = Get-ChildItem -Path $RootFolder -Include "*.jpg", "*.jpeg", "*.png", "*.gif", "*.webp" -Recurse -File -ErrorAction SilentlyContinue

if ($imageFiles.Count -eq 0) {
    Write-Warning "No image files found in $RootFolder"
    Write-Host "Please check the path and ensure it contains image files." -ForegroundColor Yellow
    exit 1
}

Write-Host "Found $($imageFiles.Count) image files" -ForegroundColor Green
Write-Host ""
Write-Host "Processing images..." -ForegroundColor Yellow

$processedCount = 0
$totalFiles = $imageFiles.Count

foreach ($file in $imageFiles) {
    $processedCount++
    
    # Show progress every 100 files or at the end
    if ($processedCount % 100 -eq 0 -or $processedCount -eq $totalFiles) {
        Write-Progress -Activity "Processing Images" -Status "$processedCount of $totalFiles" -PercentComplete (($processedCount / $totalFiles) * 100)
    }
    
    try {
        # Extract lot number from filename
        $lotNumber = ""
        # FIXED: Store match result to avoid overwriting automatic $matches
        if ($file.BaseName -match "^(\d+)(?:-\d+)?$") {
            $lotMatch = $Matches[1]
            $lotNumber = $lotMatch
        } else {
            $lotNumber = $file.BaseName
        }
        
        # Get relative path for auction/category
        $relativePath = $file.Directory.FullName.Replace($RootFolder, "").TrimStart("\")
        $pathParts = $relativePath -split "\\"
        
        $auctionHouse = if ($pathParts.Count -ge 1) { $pathParts[0] } else { "Unknown" }
        $auctionName = if ($pathParts.Count -ge 2) { $pathParts[1] } else { "Unknown" }
        
        # Check if this is the main image (no -2, -3 suffix)
        $isMainImage = $file.BaseName -notmatch "-\d+$"
        
        # Count total images for this lot
        $lotPattern = "$lotNumber*"
        $totalImages = @(Get-ChildItem -Path $file.Directory -Filter "$lotPattern.*" -ErrorAction SilentlyContinue | 
            Where-Object { $_.BaseName -match "^$([regex]::Escape($lotNumber))(-\d+)?$" }).Count
        
        # Only add main images to avoid duplicates
        if ($isMainImage) {
            $inventory += [PSCustomObject]@{
                LotNumber = $lotNumber
                AuctionHouse = $auctionHouse
                AuctionName = $auctionName
                ImageCount = $totalImages
                FilePath = $file.FullName
                FileSize = [math]::Round($file.Length / 1KB, 2)
                DateDownloaded = $file.CreationTime.ToString("yyyy-MM-dd")
                MainImage = $file.Name
            }
        }
    }
    catch {
        $errorCount++
        Write-Warning "Error processing file: $($file.FullName) - $_"
    }
}

Write-Progress -Activity "Processing Images" -Completed

# FIXED: Corrected Sort-Object syntax - property names should come before the switch
$inventory = $inventory | Sort-Object -Property DateDownloaded, LotNumber -Descending

# Export to CSV
$outputPath = Join-Path $RootFolder $OutputFile

try {
    $inventory | Export-Csv -Path $outputPath -NoTypeInformation
    Write-Host ""
    Write-Host "Inventory saved to: $outputPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to save inventory: $_"
    exit 1
}

# Display summary
Write-Host ""
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "           INVENTORY SUMMARY            " -ForegroundColor Cyan
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "Statistics:" -ForegroundColor Yellow
Write-Host "  Total lots in inventory: $($inventory.Count)" -ForegroundColor White
Write-Host "  Total image files: $($imageFiles.Count)" -ForegroundColor White

if ($errorCount -gt 0) {
    Write-Host "  Processing errors: $errorCount" -ForegroundColor Red
}

# Group by auction house
$byAuctionHouse = $inventory | Group-Object AuctionHouse | Sort-Object Count -Descending
Write-Host ""
Write-Host "By Auction House:" -ForegroundColor Yellow
foreach ($group in $byAuctionHouse) {
    $percentage = [math]::Round(($group.Count / $inventory.Count) * 100, 1)
    Write-Host ("  {0,-30} {1,5} lots ({2}%)" -f $group.Name, $group.Count, $percentage) -ForegroundColor White
}

# Group by auction name for top auction houses
if ($byAuctionHouse.Count -gt 0) {
    $topHouse = $byAuctionHouse[0].Name
    $topHouseAuctions = $inventory | Where-Object { $_.AuctionHouse -eq $topHouse } | 
        Group-Object AuctionName | Sort-Object Count -Descending | Select-Object -First 5
    
    if ($topHouseAuctions.Count -gt 0) {
        Write-Host ""
        Write-Host "Top auctions in '$topHouse':" -ForegroundColor Yellow
        foreach ($auction in $topHouseAuctions) {
            Write-Host ("  {0,-30} {1,5} lots" -f $auction.Name, $auction.Count) -ForegroundColor Gray
        }
    }
}

# Calculate storage statistics
$totalSize = ($imageFiles | Measure-Object Length -Sum).Sum
$totalSizeMB = [math]::Round($totalSize / 1MB, 2)
$totalSizeGB = [math]::Round($totalSize / 1GB, 2)

Write-Host ""
Write-Host "Storage Usage:" -ForegroundColor Yellow
if ($totalSizeGB -ge 1) {
    Write-Host "  Total: $totalSizeGB GB ($totalSizeMB MB)" -ForegroundColor White
} else {
    Write-Host "  Total: $totalSizeMB MB" -ForegroundColor White
}

if ($inventory.Count -gt 0) {
    $avgSizePerLot = [math]::Round($totalSizeMB / $inventory.Count, 2)
    Write-Host "  Average per lot: $avgSizePerLot MB" -ForegroundColor White
}

# Date range
if ($inventory.Count -gt 0) {
    $dateRange = $inventory | Measure-Object -Property DateDownloaded -Minimum -Maximum
    Write-Host ""
    Write-Host "Date Range:" -ForegroundColor Yellow
    Write-Host "  Earliest: $($dateRange.Minimum)" -ForegroundColor White
    Write-Host "  Latest: $($dateRange.Maximum)" -ForegroundColor White
}

Write-Host ""
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Offer to open the CSV
$open = Read-Host "Would you like to open the inventory file? (y/n)"
if ($open -eq 'y') {
    try {
        Start-Process $outputPath
    }
    catch {
        Write-Warning "Could not open file: $_"
        Write-Host "File location: $outputPath" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Inventory generation complete!" -ForegroundColor Green
