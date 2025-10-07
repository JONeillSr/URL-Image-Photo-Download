# Analyze-AuctionDownloadLogs.ps1
# Analyzes download logs to provide insights into your auction activity

param(
    [Parameter(Mandatory=$false)]
    [string]$LogFolder,
    
    [Parameter(Mandatory=$false)]
    [int]$LastDays = 30,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToCSV
)

# If no log folder specified, look for common locations
if (-not $LogFolder) {
    $possiblePaths = @(
        ".\logs",
        ".\images\logs",
        "$env:USERPROFILE\OneDrive - AwesomeWildStuff.com\Scripts\URL-Image-Photo-Download\images\logs",
        "C:\Auctions\logs"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $LogFolder = $path
            Write-Host "Found log folder: $LogFolder" -ForegroundColor Green
            break
        }
    }
    
    if (-not $LogFolder) {
        Write-Error "Could not find log folder. Please specify with -LogFolder parameter"
        exit 1
    }
}

Write-Host "Auction Download Log Analyzer" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host ""

# Get all log files
$cutoffDate = (Get-Date).AddDays(-$LastDays)
$logFiles = Get-ChildItem -Path $LogFolder -Filter "download_log_*.txt" | 
    Where-Object { $_.CreationTime -ge $cutoffDate } | 
    Sort-Object CreationTime -Descending

if ($logFiles.Count -eq 0) {
    Write-Host "No log files found in the last $LastDays days" -ForegroundColor Yellow
    exit
}

Write-Host "Found $($logFiles.Count) log files from the last $LastDays days" -ForegroundColor Green
Write-Host ""

# Initialize statistics
$stats = @{
    TotalRuns = 0
    TotalURLsProcessed = 0
    TotalImagesDownloaded = 0
    TotalDataDownloaded = 0
    TotalLots = @()
    AuctionSites = @{}
    DailyActivity = @{}
    Errors = @()
    ProcessingTimes = @()
}

# Parse each log file
foreach ($logFile in $logFiles) {
    Write-Host "Analyzing: $($logFile.Name)" -ForegroundColor Gray
    
    $content = Get-Content $logFile.FullName -Raw
    $stats.TotalRuns++
    
    # Extract statistics using regex
    if ($content -match "URLs Processed:\s+(\d+)\s+/\s+(\d+)") {
        $stats.TotalURLsProcessed += [int]$matches[1]
    }
    
    if ($content -match "Images Downloaded:\s+(\d+)") {
        $stats.TotalImagesDownloaded += [int]$matches[1]
    }
    
    if ($content -match "Total Data Downloaded:\s+([\d.]+)\s+(\w+)") {
        $size = [double]$matches[1]
        $unit = $matches[2]
        
        switch ($unit) {
            "GB" { $stats.TotalDataDownloaded += $size * 1GB }
            "MB" { $stats.TotalDataDownloaded += $size * 1MB }
            "KB" { $stats.TotalDataDownloaded += $size * 1KB }
            default { $stats.TotalDataDownloaded += $size }
        }
    }
    
    # Extract lot numbers
    if ($content -match "Lot Numbers Processed:\s*\n(.+?)(?:\n={10}|\z)") {
        $lotLine = $matches[1].Trim()
        if ($lotLine -and $lotLine -ne "") {
            $lots = $lotLine -split ",\s*"
            $stats.TotalLots += $lots
        }
    }
    
    # Extract auction sites
    $urlMatches = [regex]::Matches($content, "Processing URL.*?https?://([^/]+)")
    foreach ($match in $urlMatches) {
        $domain = $match.Groups[1].Value
        if (-not $stats.AuctionSites.ContainsKey($domain)) {
            $stats.AuctionSites[$domain] = 0
        }
        $stats.AuctionSites[$domain]++
    }
    
    # Track daily activity
    $date = $logFile.CreationTime.Date.ToString("yyyy-MM-dd")
    if (-not $stats.DailyActivity.ContainsKey($date)) {
        $stats.DailyActivity[$date] = @{
            Runs = 0
            Images = 0
        }
    }
    $stats.DailyActivity[$date].Runs++
    
    if ($content -match "Images Downloaded:\s+(\d+)") {
        $stats.DailyActivity[$date].Images += [int]$matches[1]
    }
    
    # Extract processing time
    if ($content -match "Duration:\s+([\d:]+)") {
        $stats.ProcessingTimes += $matches[1]
    }
    
    # Extract errors
    $errorMatches = [regex]::Matches($content, "\[Error\] (.+)")
    foreach ($match in $errorMatches) {
        $stats.Errors += @{
            File = $logFile.Name
            Error = $match.Groups[1].Value
        }
    }
}

# Display results
Write-Host ""
Write-Host "=== SUMMARY STATISTICS ===" -ForegroundColor Cyan

Write-Host ""
Write-Host "Activity Overview:" -ForegroundColor Yellow
Write-Host "  Total script runs: $($stats.TotalRuns)"
Write-Host "  Total URLs processed: $($stats.TotalURLsProcessed)"
Write-Host "  Total images downloaded: $($stats.TotalImagesDownloaded)"
$totalDataGB = [math]::Round($stats.TotalDataDownloaded / 1GB, 2)
Write-Host "  Total data downloaded: $totalDataGB GB"
Write-Host "  Unique lot numbers: $($stats.TotalLots | Select-Object -Unique | Measure-Object | Select-Object -ExpandProperty Count)"

Write-Host ""
Write-Host "Auction Sites Used:" -ForegroundColor Yellow
$stats.AuctionSites.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    Write-Host "  $($_.Key): $($_.Value) lots"
}

Write-Host ""
Write-Host "Daily Activity (Last 7 Days):" -ForegroundColor Yellow
$recentDays = $stats.DailyActivity.GetEnumerator() | 
    Sort-Object Name -Descending | 
    Select-Object -First 7

foreach ($day in $recentDays) {
    $dayOfWeek = [datetime]::Parse($day.Name).ToString("ddd")
    Write-Host "  $($day.Name) ($dayOfWeek): $($day.Value.Runs) runs, $($day.Value.Images) images"
}

# Average processing time
if ($stats.ProcessingTimes.Count -gt 0) {
    Write-Host ""
    Write-Host "Performance Metrics:" -ForegroundColor Yellow
    
    # Convert times to seconds for averaging
    $totalSeconds = 0
    foreach ($time in $stats.ProcessingTimes) {
        $parts = $time -split ":"
        $totalSeconds += [int]$parts[0] * 3600 + [int]$parts[1] * 60 + [int]$parts[2]
    }
    
    $avgSeconds = $totalSeconds / $stats.ProcessingTimes.Count
    $avgTime = [TimeSpan]::FromSeconds($avgSeconds)
    Write-Host "  Average processing time: $($avgTime.ToString("hh\:mm\:ss"))"
    Write-Host "  Average images per run: $([math]::Round($stats.TotalImagesDownloaded / $stats.TotalRuns, 1))"
}

# Recent errors
if ($stats.Errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Recent Errors (Last 5):" -ForegroundColor Yellow
    $stats.Errors | Select-Object -Last 5 | ForEach-Object {
        Write-Host "  [$($_.File)] $($_.Error)" -ForegroundColor Red
    }
}

# Export to CSV if requested
if ($ExportToCSV) {
    $exportData = @()
    
    # Create daily summary
    foreach ($day in $stats.DailyActivity.GetEnumerator()) {
        $exportData += [PSCustomObject]@{
            Date = $day.Name
            Runs = $day.Value.Runs
            ImagesDownloaded = $day.Value.Images
        }
    }
    
    $csvPath = Join-Path $LogFolder "auction_activity_summary_$(Get-Date -Format 'yyyyMMdd').csv"
    $exportData | Export-Csv -Path $csvPath -NoTypeInformation
    
    Write-Host ""
    Write-Host "Summary exported to: $csvPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "Analysis complete!" -ForegroundColor Green