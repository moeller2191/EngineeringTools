# Create Excel workbook from CSV files
param(
    [string]$CsvDir = "C:\Scripts\EngineeringTools\ExportedTables",
    [string]$ExcelPath = "C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"
)

Write-Host "Creating Excel workbook from CSV files..." -ForegroundColor Green

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Create new workbook
    $workbook = $excel.Workbooks.Add()
    
    # Remove default worksheets except first
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }
    
    # Get CSV files
    $csvFiles = Get-ChildItem -Path $CsvDir -Filter "*.csv"
    Write-Host "Found $($csvFiles.Count) CSV files to import" -ForegroundColor Yellow
    
    $worksheetIndex = 1
    
    foreach ($csvFile in $csvFiles) {
        $tableName = $csvFile.BaseName
        Write-Host "Importing: $tableName" -ForegroundColor Yellow
        
        # Get or create worksheet
        if ($worksheetIndex -eq 1) {
            $worksheet = $workbook.Worksheets.Item(1)
        } else {
            $worksheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $workbook.Worksheets.Item($workbook.Worksheets.Count))
        }
        
        # Set worksheet name (truncate if too long for Excel)
        if ($tableName.Length -gt 31) {
            $worksheet.Name = $tableName.Substring(0, 28) + "..."
        } else {
            $worksheet.Name = $tableName
        }
        
        # Import CSV data
        $queryTable = $worksheet.QueryTables.Add("TEXT;$($csvFile.FullName)", $worksheet.Range("A1"))
        $queryTable.TextFileParseType = 1  # xlDelimited
        $queryTable.TextFileCommaDelimiter = $true
        $queryTable.TextFileConsecutiveDelimiter = $false
        $queryTable.TextFileTabDelimiter = $false
        $queryTable.TextFileSemicolonDelimiter = $false
        $queryTable.TextFileSpaceDelimiter = $false
        $queryTable.TextFileTextQualifier = 1  # xlTextQualifierDoubleQuote
        $queryTable.Refresh()
        
        # Format headers
        $headerRange = $worksheet.Range("A1").EntireRow
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = 15123099  # Light blue
        
        # Auto-fit columns
        $worksheet.Columns.AutoFit()
        
        # Delete query table (keep data, remove connection)
        $queryTable.Delete()
        
        Write-Host "  Imported successfully" -ForegroundColor Green
        $worksheetIndex++
    }
    
    # Save workbook
    Write-Host "Saving Excel workbook..." -ForegroundColor Yellow
    if (Test-Path $ExcelPath) {
        Remove-Item $ExcelPath -Force
    }
    
    $workbook.SaveAs($ExcelPath, 51) # xlOpenXMLWorkbook
    Write-Host "Excel workbook saved: $ExcelPath" -ForegroundColor Green
    
    # Show summary
    Write-Host "`nWorkbook Summary:" -ForegroundColor Blue
    foreach ($ws in $workbook.Worksheets) {
        $usedRange = $ws.UsedRange
        if ($usedRange) {
            $rowCount = $usedRange.Rows.Count - 1  # Subtract header row
            Write-Host "  $($ws.Name): $rowCount rows" -ForegroundColor Cyan
        }
    }
    
    # Cleanup
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "`nMigration completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "Error creating Excel workbook: $($_.Exception.Message)" -ForegroundColor Red
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}