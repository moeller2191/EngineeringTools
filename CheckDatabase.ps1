# Check database population
try {
    Add-Type -Path "C:\Scripts\EngineeringTools\XMLIndexer\bin\Debug\net6.0\Microsoft.Data.Sqlite.dll"
    
    $connectionString = "Data Source=C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"
    $connection = New-Object Microsoft.Data.Sqlite.SqliteConnection($connectionString)
    $connection.Open()
    
    # Check SalesOrders table
    $salesCmd = New-Object Microsoft.Data.Sqlite.SqliteCommand("SELECT COUNT(*) FROM SalesOrders", $connection)
    $salesCount = $salesCmd.ExecuteScalar()
    Write-Host "Sales Orders in database: $salesCount" -ForegroundColor Green
    
    # Check ProgrammedParts table  
    $partsCmd = New-Object Microsoft.Data.Sqlite.SqliteCommand("SELECT COUNT(*) FROM ProgrammedParts", $connection)
    $partsCount = $partsCmd.ExecuteScalar()
    Write-Host "Programmed Parts in database: $partsCount" -ForegroundColor Green
    
    # Show sample data
    if ($salesCount -gt 0) {
        Write-Host "`nSample Sales Orders:" -ForegroundColor Yellow
        $sampleSalesCmd = New-Object Microsoft.Data.Sqlite.SqliteCommand("SELECT SalesOrder FROM SalesOrders LIMIT 5", $connection)
        $sampleSalesReader = $sampleSalesCmd.ExecuteReader()
        while ($sampleSalesReader.Read()) {
            Write-Host "  - $($sampleSalesReader['SalesOrder'])" -ForegroundColor Cyan
        }
        $sampleSalesReader.Close()
    }
    
    if ($partsCount -gt 0) {
        Write-Host "`nSample Programmed Parts:" -ForegroundColor Yellow
        $samplePartsCmd = New-Object Microsoft.Data.Sqlite.SqliteCommand("SELECT PartNumber FROM ProgrammedParts LIMIT 5", $connection)
        $samplePartsReader = $samplePartsCmd.ExecuteReader()
        while ($samplePartsReader.Read()) {
            Write-Host "  - $($samplePartsReader['PartNumber'])" -ForegroundColor Cyan
        }
        $samplePartsReader.Close()
    }
    
    $connection.Close()
}
catch {
    Write-Host "Error checking database: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}