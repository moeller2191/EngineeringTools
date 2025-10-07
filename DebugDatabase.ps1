# Debug script to check XMLIndex database
$dbPath = "c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"

Write-Host "Checking XMLIndex database at: $dbPath"

# Using .NET SQLite to query the database
Add-Type -AssemblyName System.Data
Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Data.Sqlite.7.0.12\lib\net6.0\Microsoft.Data.Sqlite.dll"

try {
    $connectionString = "Data Source=$dbPath"
    $connection = New-Object Microsoft.Data.Sqlite.SqliteConnection($connectionString)
    $connection.Open()
    
    Write-Host "`n=== DATABASE TABLES ==="
    $cmd = $connection.CreateCommand()
    $cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table'"
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        Write-Host "Table: $($reader['name'])"
    }
    $reader.Close()
    
    Write-Host "`n=== XMLFILES COUNT ==="
    $cmd.CommandText = "SELECT COUNT(*) as count FROM XMLFiles"
    $reader = $cmd.ExecuteReader()
    if ($reader.Read()) {
        Write-Host "XMLFiles count: $($reader['count'])"
    }
    $reader.Close()
    
    Write-Host "`n=== SAMPLE XMLFILES ==="
    $cmd.CommandText = "SELECT PartNumber, Revision, Release, FileName FROM XMLFiles LIMIT 10"
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        Write-Host "Part: $($reader['PartNumber']), Rev: $($reader['Revision']), Rel: $($reader['Release']), File: $($reader['FileName'])"
    }
    $reader.Close()
    
    Write-Host "`n=== COMPONENTS COUNT ==="
    $cmd.CommandText = "SELECT COUNT(*) as count FROM Components"
    $reader = $cmd.ExecuteReader()
    if ($reader.Read()) {
        Write-Host "Components count: $($reader['count'])"
    }
    $reader.Close()
    
    Write-Host "`n=== MRP TABLE CHECK ==="
    $cmd.CommandText = "SELECT COUNT(*) as count FROM MrpPriorityList"
    $reader = $cmd.ExecuteReader()
    if ($reader.Read()) {
        Write-Host "MrpPriorityList count: $($reader['count'])"
    }
    $reader.Close()
    
    $connection.Close()
    Write-Host "`nDatabase check complete!"
    
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    Write-Host "Make sure Microsoft.Data.Sqlite is available"
}