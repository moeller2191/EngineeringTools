# Simple SQLite query test
Add-Type -AssemblyName System.Data

$connectionString = "Data Source=c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"

try {
    # Check if we can access SQLite directly
    $assembly = [Reflection.Assembly]::LoadFrom("c:\Scripts\EngineeringTools\XMLIndexer\bin\Debug\net6.0\Microsoft.Data.Sqlite.dll")
    
    $connectionType = $assembly.GetType("Microsoft.Data.Sqlite.SqliteConnection")
    $commandType = $assembly.GetType("Microsoft.Data.Sqlite.SqliteCommand")
    
    $connection = [Activator]::CreateInstance($connectionType, $connectionString)
    $connection.Open()
    
    # Check SalesOrders
    $cmd = [Activator]::CreateInstance($commandType, "SELECT COUNT(*) FROM SalesOrders", $connection)
    $salesOrderCount = $cmd.ExecuteScalar()
    Write-Host "Sales Orders: $salesOrderCount" -ForegroundColor Green
    
    # Check ProgrammedParts
    $cmd.CommandText = "SELECT COUNT(*) FROM ProgrammedParts"
    $programmedCount = $cmd.ExecuteScalar()
    Write-Host "Programmed Parts: $programmedCount" -ForegroundColor Green
    
    # Show samples if any data
    if ($salesOrderCount -gt 0) {
        $cmd.CommandText = "SELECT SalesOrder FROM SalesOrders LIMIT 5"
        $reader = $cmd.ExecuteReader()
        Write-Host "`nSample Sales Orders:" -ForegroundColor Cyan
        while ($reader.Read()) {
            Write-Host "  - $($reader.GetString(0))"
        }
        $reader.Close()
    }
    
    if ($programmedCount -gt 0) {
        $cmd.CommandText = "SELECT PartNumber FROM ProgrammedParts LIMIT 5" 
        $reader = $cmd.ExecuteReader()
        Write-Host "`nSample Programmed Parts:" -ForegroundColor Cyan
        while ($reader.Read()) {
            Write-Host "  - $($reader.GetString(0))"
        }
        $reader.Close()
    }
    
    $connection.Close()
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}