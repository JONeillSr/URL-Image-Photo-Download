# Generate-DashboardData.ps1
# Generates JSON data file for the web dashboard

param(
    [Parameter(Mandatory=$false)]
    [string]$LogFolder,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\dashboard_data.json",
    
    [Parameter(Mandatory=$false)]
    [int]$LastDays = 90
)

# Find log folder if not specified
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
}

if (-not $LogFolder -or -not (Test-Path $LogFolder)) {
    Write-Error "Log folder not found. Please specify with -LogFolder parameter"
    exit 1
}

Write-Host "Generating Dashboard Data..." -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan

# Get log files
$cutoffDate = (Get-Date).AddDays(-$LastDays)
$logFiles = Get-ChildItem -Path $LogFolder -Filter "download_log_*.txt" | 
    Where-Object { $_.CreationTime -ge $cutoffDate } | 
    Sort-Object CreationTime

if ($logFiles.Count -eq 0) {
    Write-Warning "No log files found in the last $LastDays days"
    
    # Create empty data structure
    $dashboardData = @{
        generated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        logCount = 0
        stats = @{}
    }
    
    $dashboardData | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
    exit
}

Write-Host "Processing $($logFiles.Count) log files..." -ForegroundColor Yellow

# Initialize data structure
$dashboardData = @{
    generated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    logCount = $logFiles.Count
    stats = @{
        totalRuns = 0
        totalUrlsProcessed = 0
        totalUrlsTotal = 0
        totalImagesFound = 0
        totalImagesDownloaded = 0
        totalImagesSkipped = 0
        totalImagesFailed = 0
        totalDataBytes = 0
        uniqueLots = @()
        auctionSites = @{}
        dailyActivity = @{}
        errors = @()
        processingTimes = @()
        timeline = @{}
        recentLogs = @()
    }
}

# Process each log file
foreach ($logFile in $logFiles) {
    Write-Host "  Processing: $($logFile.Name)" -ForegroundColor Gray
    
    $content = Get-Content $logFile.FullName -Raw
    $dashboardData.stats.totalRuns++
    
    # Create log entry
    $logEntry = @{
        fileName = $logFile.Name
        date = $logFile.CreationTime.ToString("yyyy-MM-dd")
        time = $logFile.CreationTime.ToString("HH:mm:ss")
    }
    
    # Extract URLs processed
    if ($content -match "URLs Processed:\s+(\d+)\s+\/\s+(\d+)") {
        $dashboardData.stats.totalUrlsProcessed += [int]$matches[1]
        $dashboardData.stats.totalUrlsTotal += [int]$matches[2]
        $logEntry.urlsProcessed = [int]$matches[1]
        $logEntry.urlsTotal = [int]$matches[2]
    }
    
    # Extract image statistics
    if ($content -match "Images Found:\s+(\d+)") {
        $dashboardData.stats.totalImagesFound += [int]$matches[1]
        $logEntry.imagesFound = [int]$matches[1]
    }
    
    if ($content -match "Images Downloaded:\s+(\d+)") {
        $dashboardData.stats.totalImagesDownloaded += [int]$matches[1]
        $logEntry.imagesDownloaded = [int]$matches[1]
    }
    
    if ($content -match "Images Skipped[^:]*:\s+(\d+)") {
        $dashboardData.stats.totalImagesSkipped += [int]$matches[1]
        $logEntry.imagesSkipped = [int]$matches[1]
    }
    
    if ($content -match "Images Failed:\s+(\d+)") {
        $dashboardData.stats.totalImagesFailed += [int]$matches[1]
        $logEntry.imagesFailed = [int]$matches[1]
    }
    
    # Extract data size
    if ($content -match "Total Data Downloaded:\s+([\d.]+)\s+(\w+)") {
        $size = [double]$matches[1]
        $unit = $matches[2]
        $bytes = 0
        
        switch ($unit) {
            "GB" { $bytes = $size * 1GB }
            "MB" { $bytes = $size * 1MB }
            "KB" { $bytes = $size * 1KB }
            default { $bytes = $size }
        }
        
        $dashboardData.stats.totalDataBytes += $bytes
        $logEntry.dataBytes = $bytes
    }
    
    # Extract lot numbers
    if ($content -match "Lot Numbers Processed:\s*\n(.+?)(?:\n={10}|\z)") {
        $lotLine = $matches[1].Trim()
        if ($lotLine -and $lotLine -ne "") {
            $lots = $lotLine -split ",\s*" | Where-Object { $_ -ne "" }
            foreach ($lot in $lots) {
                if ($dashboardData.stats.uniqueLots -notcontains $lot) {
                    $dashboardData.stats.uniqueLots += $lot
                }
            }
            $logEntry.lotNumbers = $lots
        }
    }
    
    # Extract auction sites
    $urlMatches = [regex]::Matches($content, "Processing URL.*?https?://([^/]+)")
    $logEntry.sites = @{}
    foreach ($match in $urlMatches) {
        $domain = $match.Groups[1].Value
        
        if (-not $dashboardData.stats.auctionSites.ContainsKey($domain)) {
            $dashboardData.stats.auctionSites[$domain] = 0
        }
        $dashboardData.stats.auctionSites[$domain]++
        
        if (-not $logEntry.sites.ContainsKey($domain)) {
            $logEntry.sites[$domain] = 0
        }
        $logEntry.sites[$domain]++
    }
    
    # Daily activity
    $date = $logFile.CreationTime.Date.ToString("yyyy-MM-dd")
    if (-not $dashboardData.stats.dailyActivity.ContainsKey($date)) {
        $dashboardData.stats.dailyActivity[$date] = @{
            runs = 0
            images = 0
            data = 0
        }
    }
    $dashboardData.stats.dailyActivity[$date].runs++
    
    if ($logEntry.imagesDownloaded) {
        $dashboardData.stats.dailyActivity[$date].images += $logEntry.imagesDownloaded
    }
    if ($logEntry.dataBytes) {
        $dashboardData.stats.dailyActivity[$date].data += $logEntry.dataBytes
    }
    
    # Timeline (cumulative)
    if (-not $dashboardData.stats.timeline.ContainsKey($date)) {
        $dashboardData.stats.timeline[$date] = 0
    }
    if ($logEntry.imagesDownloaded) {
        $dashboardData.stats.timeline[$date] += $logEntry.imagesDownloaded
    }
    
    # Processing time
    if ($content -match "Duration:\s+([\d:]+)") {
        $dashboardData.stats.processingTimes += $matches[1]
        $logEntry.duration = $matches[1]
    }
    
    # Errors
    $errorMatches = [regex]::Matches($content, "\[Error\] (.+)")
    foreach ($match in $errorMatches) {
        $dashboardData.stats.errors += @{
            file = $logFile.Name
            date = $date
            error = $match.Groups[1].Value.Substring(0, [Math]::Min(200, $match.Groups[1].Value.Length))
        }
    }
    
    # Add to recent logs
    $dashboardData.stats.recentLogs += $logEntry
}

# Sort and limit data
$dashboardData.stats.recentLogs = $dashboardData.stats.recentLogs | 
    Sort-Object { [DateTime]::Parse("$($_.date) $($_.time)") } -Descending |
    Select-Object -First 100

$dashboardData.stats.errors = $dashboardData.stats.errors | 
    Sort-Object date -Descending | 
    Select-Object -First 50

$dashboardData.stats.uniqueLots = $dashboardData.stats.uniqueLots | 
    Select-Object -Unique

# Calculate summary statistics
$dashboardData.stats.successRate = if ($dashboardData.stats.totalUrlsTotal -gt 0) {
    [Math]::Round(($dashboardData.stats.totalUrlsProcessed / $dashboardData.stats.totalUrlsTotal) * 100, 2)
} else { 0 }

$dashboardData.stats.avgImagesPerRun = if ($dashboardData.stats.totalRuns -gt 0) {
    [Math]::Round($dashboardData.stats.totalImagesDownloaded / $dashboardData.stats.totalRuns, 2)
} else { 0 }

# Convert to JSON and save
Write-Host "Saving dashboard data to: $OutputPath" -ForegroundColor Green
$json = $dashboardData | ConvertTo-Json -Depth 10 -Compress
$json | Out-File -FilePath $OutputPath -Encoding UTF8

# Also create a non-compressed version for readability
$readablePath = $OutputPath -replace '\.json$', '_readable.json'
$dashboardData | ConvertTo-Json -Depth 10 | Out-File -FilePath $readablePath -Encoding UTF8

Write-Host "Dashboard data generated successfully!" -ForegroundColor Green
Write-Host "  Compressed: $OutputPath" -ForegroundColor Gray
Write-Host "  Readable: $readablePath" -ForegroundColor Gray
Write-Host ""
Write-Host "Open the dashboard HTML file in your browser and load this JSON file" -ForegroundColor Yellow