# Enhanced PowerShell script for Access to Excel migration with error handling
param(
    [string]$AccessDbPath = "C:\Scripts\EngineeringTools\JobNoBurnt.accdb",
    [string]$ExcelFilePath = "C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"
)

Write-Host "Starting Access to Excel migration..." -ForegroundColor Green

try {
    # Try different connection strings
    $connectionStrings = @(
        "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$AccessDbPath;",
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessDbPath;",
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$AccessDbPath;"
    )
    
    $conn = $null
    $connectionString = $null
    
    foreach ($connStr in $connectionStrings) {
        try {
            Write-Host "Trying connection string: $connStr" -ForegroundColor Yellow
            $conn = New-Object System.Data.OleDb.OleDbConnection($connStr)
            $conn.Open()
            $connectionString = $connStr
            Write-Host "Successfully connected!" -ForegroundColor Green
            break
        }
        catch {
            Write-Host "Failed with this connection string: $($_.Exception.Message)" -ForegroundColor Red
            if ($conn) { $conn.Dispose() }
            $conn = $null
        }
    }
    
    if (-not $conn) {
        throw "Could not establish connection to Access database with any provider"
    }
    
    # Get table names
    Write-Host "Retrieving table list..." -ForegroundColor Yellow
    $tables = $conn.GetSchema("Tables")
    $userTables = $tables | Where-Object { $_["TABLE_TYPE"] -eq "TABLE" -and $_["TABLE_NAME"] -notlike "MSys*" }
    
    Write-Host "Found $($userTables.Count) user tables:" -ForegroundColor Green
    foreach ($table in $userTables) {
        Write-Host "  - $($table['TABLE_NAME'])" -ForegroundColor Cyan
    }
    
    # Create Excel application
    Write-Host "Creating Excel application..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    
    # Create new workbook
    $workbook = $excel.Workbooks.Add()
    
    # Remove extra worksheets, keep first one
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }
    
    $worksheetIndex = 1
    
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
            
            # Get or create worksheet
            if ($worksheetIndex -eq 1) {
                $worksheet = $workbook.Worksheets.Item(1)
                $worksheet.Name = $tableName
            } else {
                $worksheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item($workbook.Worksheets.Count))
                $worksheet.Name = $tableName
            }
            
            if ($dataTable.Rows.Count -gt 0) {
                # Write headers
                for ($col = 0; $col -lt $dataTable.Columns.Count; $col++) {
                    $worksheet.Cells.Item(1, $col + 1) = $dataTable.Columns[$col].ColumnName
                    $worksheet.Cells.Item(1, $col + 1).Font.Bold = $true
                }
                
                # Write data
                for ($row = 0; $row -lt $dataTable.Rows.Count; $row++) {
                    for ($col = 0; $col -lt $dataTable.Columns.Count; $col++) {
                        $value = $dataTable.Rows[$row][$col]
                        if ($value -ne [System.DBNull]::Value) {
                            $worksheet.Cells.Item($row + 2, $col + 1) = $value
                        }
                    }
                }
                
                # Auto-fit columns
                $worksheet.Columns.AutoFit()
                Write-Host "  Successfully migrated $($dataTable.Rows.Count) records" -ForegroundColor Green
            } else {
                Write-Host "  Table is empty" -ForegroundColor Yellow
            }
            
            $worksheetIndex++
        }
        catch {
            Write-Host "  Error processing table $tableName`: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Save workbook
    Write-Host "Saving workbook to: $ExcelFilePath" -ForegroundColor Yellow
    if (Test-Path $ExcelFilePath) {
        Remove-Item $ExcelFilePath -Force
    }
    
    $workbook.SaveAs($ExcelFilePath, 51) # xlOpenXMLWorkbook format
    Write-Host "Migration completed successfully!" -ForegroundColor Green
    
    # Cleanup
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
}
catch {
    Write-Host "Error during migration: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $($_.Exception)" -ForegroundColor Red
}
finally {
    if ($conn) {
        $conn.Close()
        $conn.Dispose()
    }
}

Write-Host "Script completed." -ForegroundColor Blue