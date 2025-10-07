# Quick test script to verify the data import worked
# Use .NET directly to test the database

Add-Type -Path "c:\Scripts\EngineeringTools\XMLIndexer\bin\Debug\net6.0\XMLIndexer.dll"
Add-Type -Path "c:\Scripts\EngineeringTools\XMLIndexer\bin\Debug\net6.0\Microsoft.Data.Sqlite.dll"

try {
    $dbPath = "c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"
    $mrpManager = New-Object XMLIndexer.MrpDataManager($dbPath)
    
    Write-Host "Testing data import results..." -ForegroundColor Yellow
    
    # Get counts
    $salesOrders = $mrpManager.GetCheckedSalesOrdersFromDatabase()
    $programmedParts = $mrpManager.GetProgrammedPartsFromDatabase()
    
    Write-Host "Sales Orders in database: $($salesOrders.Count)" -ForegroundColor Green
    Write-Host "Programmed Parts in database: $($programmedParts.Count)" -ForegroundColor Green
    
    if ($salesOrders.Count -gt 0) {
        Write-Host "`nFirst 5 Sales Orders:" -ForegroundColor Cyan
        $salesOrders[0..4] | ForEach-Object { Write-Host "  - $_" }
    }
    
    if ($programmedParts.Count -gt 0) {
        Write-Host "`nFirst 5 Programmed Parts:" -ForegroundColor Cyan
        $programmedParts[0..4] | ForEach-Object { Write-Host "  - $_" }
    }
    
    # Test specific lookups
    if ($salesOrders.Count -gt 0) {
        $testSO = $salesOrders[0]
        $found = $mrpManager.CheckSalesOrderInDatabase($testSO)
        Write-Host "`nTest lookup for Sales Order '$testSO': $found" -ForegroundColor Yellow
    }
    
    if ($programmedParts.Count -gt 0) {
        $testPart = $programmedParts[0]
        $found = $mrpManager.CheckPartProgrammedInDatabase($testPart)
        Write-Host "Test lookup for Part '$testPart': $found" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}