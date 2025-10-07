# Quick test script to verify BidSpotter connectivity before running the full download

# Configure TLS
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "Testing BidSpotter Connection..." -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan

# Test URLs from your CSV
$testUrls = @(
    "https://www.bidspotter.com/en-us/auction-catalogues/new-mill-capital/catalogue-id-bscnew10468/lot-0799f80a-87e5-4935-b180-b32c013fdb39",
    "https://www.bidspotter.com/en-us/auction-catalogues/new-mill-capital/catalogue-id-bscnew10468/lot-b1c69874-b83b-4578-952c-b32c013fdb35"
)

$PSVersion = $PSVersionTable.PSVersion.Major
Write-Host "PowerShell Version: $PSVersion" -ForegroundColor Yellow

foreach ($url in $testUrls) {
    Write-Host ""
    Write-Host "Testing: $url" -ForegroundColor Yellow
    
    try {
        $response = $null
        
        if ($PSVersion -ge 7) {
            # PowerShell 7+
            Write-Host "  Using PowerShell 7+ method with AllowInsecureRedirect..." -ForegroundColor Gray
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing -AllowInsecureRedirect
        }
        else {
            # PowerShell 5.1
            Write-Host "  Using PowerShell 5.1 method..." -ForegroundColor Gray
            try {
                $response = Invoke-WebRequest -Uri $url -UseBasicParsing
            }
            catch {
                if ($_.Exception.Message -match "insecure redirection") {
                    Write-Host "  Handling redirect with WebClient..." -ForegroundColor Yellow
                    $webClient = New-Object System.Net.WebClient
                    $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                    $content = $webClient.DownloadString($url)
                    $webClient.Dispose()
                    
                    # Create a mock response object for consistency
                    $response = @{
                        StatusCode = 200
                        Content = $content
                    }
                }
                else {
                    throw
                }
            }
        }
        
        if ($response) {
            Write-Host "  ✓ Success! Status: $($response.StatusCode)" -ForegroundColor Green
            
            # Try to find images in the content
            $content = if ($response.Content) { $response.Content } else { $response }
            $imageCount = ([regex]::Matches($content, '<img[^>]+src="[^"]+"')).Count
            Write-Host "  ✓ Found $imageCount img tags in the page" -ForegroundColor Green
            
            # Look for lot-specific images
            $lotImages = ([regex]::Matches($content, 'cdn\.globalauctionplatform\.com[^"]+\.(?:jpg|jpeg|png)')).Count
            if ($lotImages -gt 0) {
                Write-Host "  ✓ Found $lotImages potential lot images from CDN" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host "  ✗ Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host "Test complete!" -ForegroundColor Green
Write-Host ""
Write-Host "If you see green checkmarks above, your system is ready to run the full script." -ForegroundColor Yellow
Write-Host "If you see red X marks, we need to troubleshoot further." -ForegroundColor Yellow