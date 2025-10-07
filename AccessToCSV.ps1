# Simple Access to CSV migration script
param(
    [string]$AccessDbPath = "C:\Scripts\EngineeringTools\JobNoBurnt.accdb",
    [string]$OutputDir = "C:\Scripts\EngineeringTools\ExportedTables"
)

Write-Host "Starting Access to CSV migration..." -ForegroundColor Green

# Create output directory
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force
    Write-Host "Created output directory: $OutputDir" -ForegroundColor Yellow
}

try {
    # Connection string
    $connectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$AccessDbPath;"
    
    Write-Host "Connecting to database..." -ForegroundColor Yellow
    $conn = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $conn.Open()
    Write-Host "Successfully connected!" -ForegroundColor Green
    
    # Get table names
    $tables = $conn.GetSchema("Tables")
    $userTables = $tables | Where-Object { $_["TABLE_TYPE"] -eq "TABLE" -and $_["TABLE_NAME"] -notlike "MSys*" }
    
    Write-Host "Found $($userTables.Count) user tables:" -ForegroundColor Green
    
    foreach ($table in $userTables) {
        $tableName = $table["TABLE_NAME"]
        Write-Host "Processing table: $tableName" -ForegroundColor Yellow
        
        try {
            # Query the table
            $command = $conn.CreateCommand()
            $command.CommandText = "SELECT * FROM [$tableName]"
            $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $rowCount = $adapter.Fill($dataTable)
            
            Write-Host "  Found $rowCount rows" -ForegroundColor Cyan
            
            # Export to CSV
            $csvPath = Join-Path $OutputDir "$tableName.csv"
            
            # Create CSV content
            $csvContent = @()
            
            # Add headers
            $headers = $dataTable.Columns | ForEach-Object { $_.ColumnName }
            $csvContent += $headers -join ","
            
            # Add data rows
            foreach ($row in $dataTable.Rows) {
                $rowData = @()
                foreach ($column in $dataTable.Columns) {
                    $value = $row[$column]
                    if ($value -eq [System.DBNull]::Value -or $value -eq $null) {
                        $rowData += '""'
                    } else {
                        # Escape quotes and wrap in quotes if contains comma
                        $valueStr = $value.ToString().Replace('"', '""')
                        if ($valueStr.Contains(",") -or $valueStr.Contains('"') -or $valueStr.Contains("`n")) {
                            $rowData += "`"$valueStr`""
                        } else {
                            $rowData += $valueStr
                        }
                    }
                }
                $csvContent += $rowData -join ","
            }
            
            # Write CSV file
            $csvContent | Out-File -FilePath $csvPath -Encoding UTF8
            Write-Host "  Exported to: $csvPath" -ForegroundColor Green
            
        } catch {
            Write-Host "  Error processing table $tableName`: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "`nMigration completed! CSV files saved in: $OutputDir" -ForegroundColor Green
    Write-Host "You can now import these CSV files into Excel or any other application." -ForegroundColor Yellow
    
} catch {
    Write-Host "Error during migration: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($conn) {
        $conn.Close()
        $conn.Dispose()
    }
}

# List exported files
Write-Host "`nExported files:" -ForegroundColor Blue
Get-ChildItem -Path $OutputDir -Filter "*.csv" | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor Cyan
}