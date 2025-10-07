# Access Database Migration Script
# Extracts data from JobNoBurnt.accdb and migrates to Excel

Write-Host "=== ACCESS DATABASE MIGRATION TO EXCEL ===" -ForegroundColor Green

# Set paths
$accessDbPath = "C:\Scripts\EngineeringTools\JobNoBurnt.accdb"
$excelOutputPath = "C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"

# Check if Access database exists
if (-not (Test-Path $accessDbPath)) {
    Write-Host "ERROR: Access database not found at $accessDbPath" -ForegroundColor Red
    exit 1
}

Write-Host "Found Access database: $accessDbPath" -ForegroundColor Green

try {
    # Create ADODB connection to Access database
    $accessConnection = New-Object -ComObject ADODB.Connection
    $accessConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessDbPath;")
    
    Write-Host "Connected to Access database" -ForegroundColor Green
    
    # Get list of tables in the database
    $schema = $accessConnection.OpenSchema(20) # adSchemaTables = 20
    $tables = @()
    
    Write-Host "`n=== DATABASE TABLES ===" -ForegroundColor Yellow
    while (-not $schema.EOF) {
        $tableName = $schema.Fields("TABLE_NAME").Value
        $tableType = $schema.Fields("TABLE_TYPE").Value
        
        # Only include user tables (not system tables)
        if ($tableType -eq "TABLE" -and -not $tableName.StartsWith("MSys")) {
            $tables += $tableName
            Write-Host "Found table: $tableName" -ForegroundColor Cyan
        }
        $schema.MoveNext()
    }
    $schema.Close()
    
    if ($tables.Count -eq 0) {
        Write-Host "No user tables found in database" -ForegroundColor Yellow
        return
    }
    
    # Create Excel application
    Write-Host "`n=== CREATING EXCEL WORKBOOK ===" -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Create new workbook
    $workbook = $excel.Workbooks.Add()
    
    # Remove default sheets except the first one
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }
    
    $sheetIndex = 1
    
    foreach ($tableName in $tables) {
        Write-Host "`nProcessing table: $tableName" -ForegroundColor Cyan
        
        # Query the table data
        $recordset = $accessConnection.Execute("SELECT * FROM [$tableName]")
        
        # Create or use worksheet
        if ($sheetIndex -eq 1) {
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = $tableName
        } else {
            $worksheet = $workbook.Worksheets.Add()
            $worksheet.Name = $tableName
        }
        
        # Check if table has data
        if ($recordset.EOF) {
            Write-Host "  Table $tableName is empty" -ForegroundColor Yellow
            $recordset.Close()
            $sheetIndex++
            continue
        }
        
        # Get field names and write headers
        $fieldCount = $recordset.Fields.Count
        $fieldNames = @()
        
        for ($i = 0; $i -lt $fieldCount; $i++) {
            $fieldName = $recordset.Fields.Item($i).Name
            $fieldNames += $fieldName
            $worksheet.Cells.Item(1, $i + 1) = $fieldName
        }
        
        # Format header row
        $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $fieldCount))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15  # Gray background
        $headerRange.Font.ColorIndex = 1       # Black text
        
        # Write data rows
        $rowIndex = 2
        $recordCount = 0
        
        while (-not $recordset.EOF) {
            for ($i = 0; $i -lt $fieldCount; $i++) {
                $value = $recordset.Fields.Item($i).Value
                if ($null -ne $value) {
                    $worksheet.Cells.Item($rowIndex, $i + 1) = $value
                }
            }
            $recordset.MoveNext()
            $rowIndex++
            $recordCount++
        }
        
        # Auto-fit columns
        $worksheet.Columns.AutoFit() | Out-Null
        
        Write-Host "  Exported $recordCount records with $fieldCount fields" -ForegroundColor Green
        
        $recordset.Close()
        $sheetIndex++
    }
    
    # Save Excel file
    Write-Host "`n=== SAVING EXCEL FILE ===" -ForegroundColor Yellow
    
    # Remove existing file if it exists
    if (Test-Path $excelOutputPath) {
        Remove-Item $excelOutputPath -Force
        Write-Host "Removed existing Excel file" -ForegroundColor Yellow
    }
    
    $workbook.SaveAs($excelOutputPath)
    Write-Host "Saved Excel database: $excelOutputPath" -ForegroundColor Green
    
    # Close Excel
    $workbook.Close()
    $excel.Quit()
    
    # Clean up COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Close Access connection
    $accessConnection.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($accessConnection) | Out-Null
    
    Write-Host "`n=== MIGRATION COMPLETE ===" -ForegroundColor Green
    Write-Host "Excel database created with $($tables.Count) worksheets:" -ForegroundColor White
    foreach ($table in $tables) {
        Write-Host "  - $table" -ForegroundColor Cyan
    }
    Write-Host "`nFile location: $excelOutputPath" -ForegroundColor White
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    
    # Clean up on error
    if ($excel) {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) } catch { }
    }
    if ($accessConnection) {
        try { $accessConnection.Close() } catch { }
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($accessConnection) } catch { }
    }
}