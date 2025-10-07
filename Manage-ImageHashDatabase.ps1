<#
.SYNOPSIS
    Manages, analyzes, and maintains the image hash database used for duplicate detection.

.DESCRIPTION
    This utility script provides comprehensive management capabilities for the SHA256 hash database
    used by the Download-ImagesFromURLs script to detect duplicate images. The database tracks
    unique images by their hash values, preventing duplicate downloads and saving bandwidth.
    
    The script supports multiple operations:
    - Show: Display database statistics and recent additions
    - Rebuild: Recreate the database from scratch by scanning all images
    - FindDuplicates: Scan for duplicate images in the file system
    - Clean: Remove database entries for missing files
    - Export: Export the database to CSV format
    - Merge: Combine databases from different folders

.PARAMETER ImageFolder
    Optional. The folder containing images and the hash database.
    If not specified, the script searches common locations:
    - .\images
    - C:\Auctions
    - User's OneDrive Scripts folder
    
.PARAMETER Action
    The operation to perform. Valid values:
    - Show (default): Display database statistics
    - Rebuild: Rebuild database from existing images
    - FindDuplicates: Find duplicate images
    - Clean: Remove entries for missing files
    - Export: Export database to CSV
    - Merge: Merge with another database

.INPUTS
    System.String
    Accepts a string path to the image folder.

.OUTPUTS
    Varies by action:
    - Show: Console output with statistics
    - Rebuild: Updated JSON database file
    - FindDuplicates: List of duplicates (optionally deleted)
    - Clean: Updated database with missing entries removed
    - Export: CSV file with database contents
    - Merge: Combined database file

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1
    
    Shows database statistics using auto-detected image folder.

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1 -ImageFolder "C:\Auctions" -Action Show
    
    Display statistics for the database in C:\Auctions.

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1 -ImageFolder "D:\Photos" -Action Rebuild
    
    Rebuild the hash database by scanning all images in D:\Photos.

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1 -Action FindDuplicates
    
    Scan for duplicate images and optionally delete them.

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1 -Action Clean
    
    Remove database entries for files that no longer exist.

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1 -Action Export
    
    Export the database to a CSV file for analysis.

.EXAMPLE
    .\Manage-ImageHashDatabase.ps1 -Action Merge
    
    Merge the current database with another folder's database.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 10/06/2025
    Version: 1.1
    Change Date: 10/06/2025
    
    The hash database is stored in a hidden .imagedb folder within the image directory.
    Database format: JSON with SHA256 hashes as keys.
    
    VERSION 1.1 CHANGES:
    - Fixed PSScriptAnalyzer warnings (renamed functions to use approved verbs)
    - Added comprehensive PowerShell help documentation
    - Fixed automatic variable $matches usage
    
    Database Structure:
    Each entry contains:
    - FilePath: Full path to the image file
    - FileSize: Size in bytes
    - DateAdded: Timestamp when added to database
    - LotNumber: Extracted lot number (if applicable)
    - SourceUrl: Original download URL (if known)

.LINK
    https://github.com/JONeillSr/Manage-ImageHashDatabase
#>

param(
    [Parameter(Mandatory=$false, Position=0, ValueFromPipeline=$true)]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_ -PathType Container)) {
            throw "Image folder path '$_' does not exist."
        }
        return $true
    })]
    [string]$ImageFolder,
    
    [Parameter(Mandatory=$false, Position=1)]
    [ValidateSet("Show", "Rebuild", "FindDuplicates", "Clean", "Export", "Merge")]
    [string]$Action = "Show"
)

# Find the image folder if not specified
if (-not $ImageFolder) {
    $possiblePaths = @(
        ".\images",
        "C:\Auctions",
        "$env:USERPROFILE\OneDrive - AwesomeWildStuff.com\Scripts\URL-Image-Photo-Download\images"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $ImageFolder = $path
            Write-Host "Found image folder: $ImageFolder" -ForegroundColor Green
            break
        }
    }
    
    if (-not $ImageFolder) {
        Write-Error "Could not find image folder. Please specify with -ImageFolder parameter"
        exit 1
    }
}

$dbPath = Join-Path $ImageFolder ".imagedb\image_hashes.json"

Write-Host "Image Hash Database Manager v1.1" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# FIXED: Using approved verb "Get" instead of "Load"
function Get-HashDatabase {
    <#
    .SYNOPSIS
        Retrieves the hash database from disk.
    
    .DESCRIPTION
        Loads the JSON hash database file and converts it to a PowerShell hashtable.
    
    .OUTPUTS
        System.Collections.Hashtable
        Returns a hashtable containing the image hash database.
    #>
    [CmdletBinding()]
    param()
    
    if (Test-Path $dbPath) {
        try {
            $json = Get-Content $dbPath -Raw
            $loaded = $json | ConvertFrom-Json
            
            $database = @{}
            foreach ($prop in $loaded.PSObject.Properties) {
                $database[$prop.Name] = @{
                    FilePath = $prop.Value.FilePath
                    FileSize = $prop.Value.FileSize
                    DateAdded = $prop.Value.DateAdded
                    LotNumber = $prop.Value.LotNumber
                    SourceUrl = $prop.Value.SourceUrl
                }
            }
            
            return $database
        }
        catch {
            Write-Error "Error loading database: $_"
            return @{}
        }
    }
    else {
        Write-Warning "No hash database found at: $dbPath"
        return @{}
    }
}

# Function to save database
function Save-HashDatabase {
    <#
    .SYNOPSIS
        Saves the hash database to disk.
    
    .DESCRIPTION
        Converts the hashtable to JSON and saves it to the .imagedb folder.
    
    .PARAMETER Database
        The hashtable containing the hash database to save.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Database
    )
    
    $dbFolder = Split-Path $dbPath -Parent
    if (-not (Test-Path $dbFolder)) {
        New-Item -ItemType Directory -Path $dbFolder -Force | Out-Null
        
        # Hide the folder on Windows
        if ($IsWindows -ne $false) {
            $folder = Get-Item $dbFolder
            $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::Hidden
        }
    }
    
    $Database | ConvertTo-Json -Depth 10 | Out-File -FilePath $dbPath -Encoding UTF8
    Write-Host "Database saved to: $dbPath" -ForegroundColor Green
}

switch ($Action) {
    "Show" {
        Write-Host "Loading hash database..." -ForegroundColor Yellow
        $db = Get-HashDatabase  # FIXED: Using renamed function
        
        if ($db.Count -eq 0) {
            Write-Host "Database is empty or not found" -ForegroundColor Red
            Write-Host "Run with -Action Rebuild to create a new database" -ForegroundColor Yellow
            exit
        }
        
        Write-Host "Database Statistics:" -ForegroundColor Cyan
        Write-Host "  Total unique images: $($db.Count)"
        
        # Calculate total size
        $totalSize = 0
        $db.Values | ForEach-Object { $totalSize += $_.FileSize }
        $totalSizeMB = [math]::Round($totalSize / 1MB, 2)
        Write-Host "  Total size tracked: $totalSizeMB MB"
        
        # Count by lot numbers
        $lotNumbers = @{}
        foreach ($entry in $db.Values) {
            if ($entry.LotNumber) {
                if ($lotNumbers.ContainsKey($entry.LotNumber)) {
                    $lotNumbers[$entry.LotNumber]++
                } else {
                    $lotNumbers[$entry.LotNumber] = 1
                }
            }
        }
        Write-Host "  Unique lot numbers: $($lotNumbers.Count)"
        
        # Find missing files
        $missingFiles = @()
        foreach ($entry in $db.Values) {
            if (-not (Test-Path $entry.FilePath)) {
                $missingFiles += $entry.FilePath
            }
        }
        
        if ($missingFiles.Count -gt 0) {
            Write-Host "  ⚠ Missing files: $($missingFiles.Count)" -ForegroundColor Yellow
        }
        
        # Recent additions
        Write-Host ""
        Write-Host "Recent Additions (Last 10):" -ForegroundColor Cyan
        $recent = $db.Values | 
            Where-Object { $_.DateAdded } |
            Sort-Object { [DateTime]::Parse($_.DateAdded) } -Descending |
            Select-Object -First 10
        
        foreach ($item in $recent) {
            $fileName = Split-Path $item.FilePath -Leaf
            $sizeKB = [math]::Round($item.FileSize / 1KB, 2)
            Write-Host "  $fileName - $sizeKB KB - Lot: $($item.LotNumber) - $($item.DateAdded)"
        }
    }
    
    "Rebuild" {
        Write-Host "Rebuilding hash database from scratch..." -ForegroundColor Yellow
        Write-Host "This will scan all images and may take several minutes." -ForegroundColor Yellow
        
        $confirm = Read-Host "Continue? (y/n)"
        if ($confirm -ne 'y') {
            Write-Host "Cancelled" -ForegroundColor Red
            exit
        }
        
        $newDb = @{}
        $imageFiles = Get-ChildItem -Path $ImageFolder -Include "*.jpg", "*.jpeg", "*.png", "*.gif", "*.webp" -Recurse -File |
            Where-Object { $_.DirectoryName -notlike "*\.imagedb*" -and $_.DirectoryName -notlike "*\logs*" }
        
        $total = $imageFiles.Count
        $current = 0
        
        Write-Host "Found $total image files to process..." -ForegroundColor Cyan
        
        foreach ($file in $imageFiles) {
            $current++
            if ($current % 10 -eq 0 -or $current -eq $total) {
                Write-Progress -Activity "Rebuilding Database" -Status "$current of $total files" -PercentComplete (($current / $total) * 100)
            }
            
            try {
                $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                
                $lotNumber = ""
                # FIXED: Store match result to avoid overwriting automatic $matches
                if ($file.BaseName -match "^(\d+)(?:-\d+)?$") {
                    $lotMatch = $Matches[1]
                    $lotNumber = $lotMatch
                }
                
                $newDb[$hash] = @{
                    FilePath = $file.FullName
                    FileSize = $file.Length
                    DateAdded = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                    LotNumber = $lotNumber
                    SourceUrl = ""
                }
            }
            catch {
                Write-Warning "Error processing $($file.FullName): $_"
            }
        }
        
        Write-Progress -Activity "Rebuilding Database" -Completed
        
        Save-HashDatabase -Database $newDb
        Write-Host "Database rebuilt with $($newDb.Count) images" -ForegroundColor Green
    }
    
    "FindDuplicates" {
        Write-Host "Scanning for duplicate images..." -ForegroundColor Yellow
        
        $imageFiles = Get-ChildItem -Path $ImageFolder -Include "*.jpg", "*.jpeg", "*.png", "*.gif", "*.webp" -Recurse -File |
            Where-Object { $_.DirectoryName -notlike "*\.imagedb*" -and $_.DirectoryName -notlike "*\logs*" }
        
        $hashes = @{}
        $duplicates = @()
        $total = $imageFiles.Count
        $current = 0
        
        Write-Host "Scanning $total files..." -ForegroundColor Cyan
        
        foreach ($file in $imageFiles) {
            $current++
            if ($current % 10 -eq 0 -or $current -eq $total) {
                Write-Progress -Activity "Scanning for Duplicates" -Status "$current of $total" -PercentComplete (($current / $total) * 100)
            }
            
            try {
                $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                
                if ($hashes.ContainsKey($hash)) {
                    $duplicates += @{
                        Original = $hashes[$hash]
                        Duplicate = $file.FullName
                        Size = $file.Length
                        Hash = $hash
                    }
                }
                else {
                    $hashes[$hash] = $file.FullName
                }
            }
            catch {
                Write-Warning "Error processing $($file.FullName): $_"
            }
        }
        
        Write-Progress -Activity "Scanning for Duplicates" -Completed
        
        if ($duplicates.Count -eq 0) {
            Write-Host "No duplicate images found!" -ForegroundColor Green
        }
        else {
            Write-Host "Found $($duplicates.Count) duplicate images:" -ForegroundColor Yellow
            
            $totalWasted = 0
            foreach ($dup in $duplicates) {
                Write-Host ""
                Write-Host "  Original:  $(Split-Path $dup.Original -Leaf)" -ForegroundColor Cyan
                Write-Host "  Duplicate: $(Split-Path $dup.Duplicate -Leaf)" -ForegroundColor Red
                Write-Host "  Size: $([math]::Round($dup.Size / 1KB, 2)) KB"
                Write-Host "  Hash: $($dup.Hash.Substring(0, 16))..."
                $totalWasted += $dup.Size
            }
            
            Write-Host ""
            Write-Host "Total space wasted: $([math]::Round($totalWasted / 1MB, 2)) MB" -ForegroundColor Yellow
            
            $delete = Read-Host "Would you like to delete the duplicates? (y/n)"
            if ($delete -eq 'y') {
                foreach ($dup in $duplicates) {
                    Remove-Item $dup.Duplicate -Force
                    Write-Host "Deleted: $(Split-Path $dup.Duplicate -Leaf)" -ForegroundColor Red
                }
                Write-Host "Deleted $($duplicates.Count) duplicate files" -ForegroundColor Green
            }
        }
    }
    
    "Clean" {
        Write-Host "Cleaning database of missing files..." -ForegroundColor Yellow
        $db = Get-HashDatabase  # FIXED: Using renamed function
        
        if ($db.Count -eq 0) {
            Write-Host "Database is empty or not found" -ForegroundColor Red
            exit
        }
        
        $toRemove = @()
        foreach ($hash in $db.Keys) {
            if (-not (Test-Path $db[$hash].FilePath)) {
                $toRemove += $hash
                Write-Host "  Missing: $($db[$hash].FilePath)" -ForegroundColor Red
            }
        }
        
        if ($toRemove.Count -eq 0) {
            Write-Host "No missing files found. Database is clean!" -ForegroundColor Green
        }
        else {
            Write-Host "Found $($toRemove.Count) missing files" -ForegroundColor Yellow
            
            $confirm = Read-Host "Remove these entries from database? (y/n)"
            if ($confirm -eq 'y') {
                foreach ($hash in $toRemove) {
                    $db.Remove($hash)
                }
                Save-HashDatabase -Database $db
                Write-Host "Removed $($toRemove.Count) entries" -ForegroundColor Green
            }
        }
    }
    
    "Export" {
        Write-Host "Exporting database to CSV..." -ForegroundColor Yellow
        $db = Get-HashDatabase  # FIXED: Using renamed function
        
        if ($db.Count -eq 0) {
            Write-Host "Database is empty or not found" -ForegroundColor Red
            exit
        }
        
        $exportData = @()
        foreach ($hash in $db.Keys) {
            $exportData += [PSCustomObject]@{
                Hash = $hash.Substring(0, [Math]::Min(16, $hash.Length)) + "..."
                FileName = Split-Path $db[$hash].FilePath -Leaf
                FilePath = $db[$hash].FilePath
                SizeKB = [math]::Round($db[$hash].FileSize / 1KB, 2)
                LotNumber = $db[$hash].LotNumber
                DateAdded = $db[$hash].DateAdded
                SourceUrl = $db[$hash].SourceUrl
            }
        }
        
        $exportPath = Join-Path $ImageFolder "hash_database_export_$(Get-Date -Format 'yyyyMMdd').csv"
        $exportData | Export-Csv -Path $exportPath -NoTypeInformation
        
        Write-Host "Database exported to: $exportPath" -ForegroundColor Green
        Write-Host "Total entries: $($exportData.Count)" -ForegroundColor Cyan
        
        $open = Read-Host "Open the exported file? (y/n)"
        if ($open -eq 'y') {
            Start-Process $exportPath
        }
    }
    
    "Merge" {
        Write-Host "Merge database with another folder's database" -ForegroundColor Yellow
        
        $otherFolder = Read-Host "Enter path to other image folder"
        if (-not (Test-Path $otherFolder)) {
            Write-Error "Folder not found: $otherFolder"
            exit
        }
        
        $otherDbPath = Join-Path $otherFolder ".imagedb\image_hashes.json"
        if (-not (Test-Path $otherDbPath)) {
            Write-Error "No database found in: $otherFolder"
            exit
        }
        
        $db1 = Get-HashDatabase  # FIXED: Using renamed function
        Write-Host "Current database: $($db1.Count) entries" -ForegroundColor Cyan
        
        # Load other database
        Write-Host "Loading other database..." -ForegroundColor Yellow
        try {
            $json = Get-Content $otherDbPath -Raw
            $loaded = $json | ConvertFrom-Json
            $db2 = @{}
            foreach ($prop in $loaded.PSObject.Properties) {
                $db2[$prop.Name] = @{
                    FilePath = $prop.Value.FilePath
                    FileSize = $prop.Value.FileSize
                    DateAdded = $prop.Value.DateAdded
                    LotNumber = $prop.Value.LotNumber
                    SourceUrl = $prop.Value.SourceUrl
                }
            }
        }
        catch {
            Write-Error "Failed to load other database: $_"
            exit
        }
        
        Write-Host "Other database: $($db2.Count) entries" -ForegroundColor Cyan
        
        $newEntries = 0
        $duplicateEntries = 0
        foreach ($hash in $db2.Keys) {
            if (-not $db1.ContainsKey($hash)) {
                $db1[$hash] = $db2[$hash]
                $newEntries++
            }
            else {
                $duplicateEntries++
            }
        }
        
        Write-Host ""
        Write-Host "Merge Results:" -ForegroundColor Green
        Write-Host "  New entries added: $newEntries" -ForegroundColor Cyan
        Write-Host "  Duplicate entries skipped: $duplicateEntries" -ForegroundColor Yellow
        Write-Host "  Total entries in merged database: $($db1.Count)" -ForegroundColor Cyan
        
        if ($newEntries -gt 0) {
            Save-HashDatabase -Database $db1
            Write-Host ""
            Write-Host "Database merge completed successfully!" -ForegroundColor Green
        }
        else {
            Write-Host ""
            Write-Host "No new entries to add. Databases already synchronized." -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "Operation complete!" -ForegroundColor Green