# Test Database State
param(
    [string]$DatabasePath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"
)

Write-Host "=== DATABASE TEST ===" -ForegroundColor Green

# Check if database file exists
if (Test-Path $DatabasePath) {
    Write-Host "Database file exists: $DatabasePath" -ForegroundColor Green
    $dbSize = (Get-Item $DatabasePath).Length
    Write-Host "Database size: $dbSize bytes" -ForegroundColor Cyan
} else {
    Write-Host "Database file NOT found: $DatabasePath" -ForegroundColor Red
    exit 1
}

# Test using .NET SQLite
Write-Host "`nTesting database connection..." -ForegroundColor Yellow

try {
    Add-Type -AssemblyName System.Data
    
    # Load SQLite assembly if available
    $sqliteAssembly = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Data.Sqlite")
    if ($sqliteAssembly) {
        Write-Host "Microsoft.Data.Sqlite loaded successfully" -ForegroundColor Green
        
        $connectionString = "Data Source=$DatabasePath"
        $connection = New-Object Microsoft.Data.Sqlite.SqliteConnection($connectionString)
        $connection.Open()
        
        # Check if MrpPriorityList table exists
        $command = $connection.CreateCommand()
        $command.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='MrpPriorityList'"
        $result = $command.ExecuteScalar()
        
        if ($result) {
            Write-Host "✅ MrpPriorityList table exists" -ForegroundColor Green
            
            # Count records
            $command.CommandText = "SELECT COUNT(*) FROM MrpPriorityList"
            $count = $command.ExecuteScalar()
            Write-Host "Records in MrpPriorityList: $count" -ForegroundColor Cyan
            
            # Check for test job
            $command.CommandText = "SELECT COUNT(*) FROM MrpPriorityList WHERE JobNumber = 'H1319-0000'"
            $testJobCount = $command.ExecuteScalar()
            Write-Host "Test job (H1319-0000) found: $testJobCount times" -ForegroundColor Cyan
            
            # Show some sample data
            $command.CommandText = "SELECT JobNumber, PartNumber, Quantity FROM MrpPriorityList LIMIT 5"
            $reader = $command.ExecuteReader()
            Write-Host "`nSample MRP data:" -ForegroundColor Yellow
            while ($reader.Read()) {
                Write-Host "  $($reader['JobNumber']) | $($reader['PartNumber']) | Qty: $($reader['Quantity'])" -ForegroundColor White
            }
            $reader.Close()
            
        } else {
            Write-Host "❌ MrpPriorityList table does NOT exist" -ForegroundColor Red
        }
        
        $connection.Close()
        
    } else {
        Write-Host "Microsoft.Data.Sqlite not available" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "Error testing database: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== TEST COMPLETE ===" -ForegroundColor Green