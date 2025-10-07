# Check what XML files are in the database for debugging
try {
    # Load System.Data.SQLite assembly
    $sqliteAssemblyPath = "C:\Scripts\EngineeringTools\packages\System.Data.SQLite.Core.1.0.116\lib\net46\System.Data.SQLite.dll"
    
    if (Test-Path $sqliteAssemblyPath) {
        Add-Type -Path $sqliteAssemblyPath -ErrorAction Stop
        
        $dbPath = "C:\Scripts\EngineeringTools\XMLIndex.db"
        $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$dbPath")
        $connection.Open()
        
        Write-Host "=== Checking XML Database Contents ===" -ForegroundColor Green
        Write-Host ""
        
        # Check total XML files
        $totalQuery = "SELECT COUNT(*) FROM XMLFiles"
        $totalCommand = New-Object System.Data.SQLite.SQLiteCommand($totalQuery, $connection)
        $totalCount = $totalCommand.ExecuteScalar()
        Write-Host "Total XML files in database: $totalCount" -ForegroundColor Yellow
        
        # Search for any part numbers containing SULL
        $searchQuery = "SELECT PartNumber, Revision, Release FROM XMLFiles WHERE PartNumber LIKE '%SULL%' LIMIT 10"
        $searchCommand = New-Object System.Data.SQLite.SQLiteCommand($searchQuery, $connection)
        $reader = $searchCommand.ExecuteReader()
        
        Write-Host ""
        Write-Host "XML files containing 'SULL':" -ForegroundColor Cyan
        $found = $false
        while ($reader.Read()) {
            $found = $true
            $partNum = $reader["PartNumber"]
            $rev = $reader["Revision"] 
            $rel = $reader["Release"]
            Write-Host "  $partNum (Rev: $rev, Rel: $rel)" -ForegroundColor White
        }
        
        if (-not $found) {
            Write-Host "  No XML files found containing 'SULL'" -ForegroundColor Red
        }
        
        $reader.Close()
        
        # Check Components table
        $compQuery = "SELECT COUNT(*) FROM Components"
        $compCommand = New-Object System.Data.SQLite.SQLiteCommand($compQuery, $connection)
        $compCount = $compCommand.ExecuteScalar()
        Write-Host ""
        Write-Host "Total components in database: $compCount" -ForegroundColor Yellow
        
        # Show some sample XML files
        Write-Host ""
        Write-Host "Sample XML files (first 10):" -ForegroundColor Cyan
        $sampleQuery = "SELECT PartNumber, Revision, Release FROM XMLFiles LIMIT 10"
        $sampleCommand = New-Object System.Data.SQLite.SQLiteCommand($sampleQuery, $connection)
        $sampleReader = $sampleCommand.ExecuteReader()
        
        while ($sampleReader.Read()) {
            $partNum = $sampleReader["PartNumber"]
            $rev = $sampleReader["Revision"] 
            $rel = $sampleReader["Release"]
            Write-Host "  $partNum (Rev: $rev, Rel: $rel)" -ForegroundColor White
        }
        
        $sampleReader.Close()
        $connection.Close()
    }
    else {
        Write-Host "SQLite assembly not found at: $sqliteAssemblyPath" -ForegroundColor Red
        Write-Host "Trying alternative method..." -ForegroundColor Yellow
        
        # Try using .NET Core approach
        cd "C:\Scripts\EngineeringTools"
        $output = dotnet run --project "XMLIndexer\XMLIndexer.csproj" -- --db-info 2>&1
        Write-Host $output
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Trying alternative: Check if database file exists..." -ForegroundColor Yellow
    $dbPath = "C:\Scripts\EngineeringTools\XMLIndex.db"
    if (Test-Path $dbPath) {
        $fileInfo = Get-Item $dbPath
        Write-Host "Database file exists: $($fileInfo.FullName)" -ForegroundColor Green
        Write-Host "File size: $($fileInfo.Length) bytes" -ForegroundColor Green
        Write-Host "Last modified: $($fileInfo.LastWriteTime)" -ForegroundColor Green
    }
    else {
        Write-Host "Database file does not exist at: $dbPath" -ForegroundColor Red
    }
}