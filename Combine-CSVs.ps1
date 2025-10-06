# Combine all CSVs and check for duplicate lot numbers
$allLots = @()
Get-ChildItem "C:\Temp\images\auction_*.csv" | ForEach-Object {
    $allLots += Import-Csv $_.FullName
}

# Group by lot number to find duplicates
$grouped = $allLots | Where-Object { $_.LotNumber -ne "" } | Group-Object LotNumber | Where-Object { $_.Count -gt 1 }
Write-Host "Duplicate lot numbers found: $($grouped.Count)"
$grouped | ForEach-Object { 
    Write-Host "  Lot $($_.Name): $($_.Count) times"
}

$noLotNumber = $allLots | Where-Object { [string]::IsNullOrWhiteSpace($_.LotNumber) }
Write-Host "Lots without lot numbers: $($noLotNumber.Count)"