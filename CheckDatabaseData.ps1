# Check database data import
param(
    [string]$DatabasePath = "c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"
)

Add-Type -AssemblyName "Microsoft.Data.Sqlite"

try {
    $connectionString = "Data Source=$DatabasePath"
    $connection = New-Object Microsoft.Data.Sqlite.SqliteConnection($connectionString)
    $connection.Open()
    
    # Check SalesOrders table
    $cmd = $connection.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM SalesOrders"
    $salesOrderCount = $cmd.ExecuteScalar()
    Write-Host "Sales Orders in database: $salesOrderCount"
    
    # Check ProgrammedParts table
    $cmd.CommandText = "SELECT COUNT(*) FROM ProgrammedParts"
    $programmedPartsCount = $cmd.ExecuteScalar()
    Write-Host "Programmed Parts in database: $programmedPartsCount"
    
    # Show sample sales orders
    if ($salesOrderCount -gt 0) {
        Write-Host "`nSample Sales Orders:"
        $cmd.CommandText = "SELECT SalesOrder FROM SalesOrders LIMIT 10"
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            Write-Host "  - $($reader['SalesOrder'])"
        }
        $reader.Close()
    }
    
    # Show sample programmed parts
    if ($programmedPartsCount -gt 0) {
        Write-Host "`nSample Programmed Parts:"
        $cmd.CommandText = "SELECT PartNumber FROM ProgrammedParts LIMIT 10"
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            Write-Host "  - $($reader['PartNumber'])"
        }
        $reader.Close()
    }
    
    $connection.Close()
}
catch {
    Write-Host "Error checking database: $_"
}