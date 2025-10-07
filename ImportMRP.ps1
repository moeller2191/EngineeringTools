# Import MRP Priority List from Excel to SQLite
param(
    [string]$ExcelPath = "C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls",
    [string]$DatabasePath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db",
    [string]$WorksheetName = "Priority List Sheet"
)

Write-Host "=== MRP PRIORITY LIST IMPORTER ===" -ForegroundColor Green

try {
    # Create Excel COM object
    Write-Host "Opening Excel file: $ExcelPath" -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    # Try to find the worksheet
    $worksheet = $null
    foreach ($ws in $workbook.Worksheets) {
        Write-Host "Found worksheet: '$($ws.Name)'" -ForegroundColor Cyan
        if ($ws.Name -like "*Priority*" -or $ws.Name -like "*Sheet*") {
            $worksheet = $ws
            $WorksheetName = $ws.Name
            break
        }
    }
    
    if (-not $worksheet) {
        # Use first worksheet if no match found
        $worksheet = $workbook.Worksheets.Item(1)
        $WorksheetName = $worksheet.Name
    }
    
    Write-Host "Using worksheet: '$WorksheetName'" -ForegroundColor Green
    
    # Find the used range
    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
    Write-Host "Data range: $rowCount rows x $colCount columns" -ForegroundColor Yellow
    
    # Read headers from first row
    Write-Host "`nColumn Headers:" -ForegroundColor Blue
    $headers = @()
    for ($col = 1; $col -le $colCount; $col++) {
        $header = $worksheet.Cells.Item(1, $col).Text.Trim()
        if ($header) {
            $headers += $header
            Write-Host "  Column $col`: $header" -ForegroundColor Cyan
        }
    }
    
    # Show sample data from first few rows
    Write-Host "`nSample Data (first 5 rows):" -ForegroundColor Blue
    for ($row = 1; $row -le [Math]::Min(6, $rowCount); $row++) {
        $rowData = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $cellValue = $worksheet.Cells.Item($row, $col).Text.Trim()
            $rowData += $cellValue
        }
        Write-Host "Row $row`: $($rowData -join ' | ')" -ForegroundColor White
    }
    
    # Export first 10 rows to CSV for analysis
    $csvPath = "C:\Scripts\EngineeringTools\MRP_Sample.csv"
    Write-Host "`nExporting sample to: $csvPath" -ForegroundColor Yellow
    
    $csvContent = @()
    for ($row = 1; $row -le [Math]::Min(10, $rowCount); $row++) {
        $rowData = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $cellValue = $worksheet.Cells.Item($row, $col).Text.Trim()
            if ($cellValue.Contains(",") -or $cellValue.Contains('"')) {
                $rowData += "`"$($cellValue.Replace('"', '""'))`""
            } else {
                $rowData += $cellValue
            }
        }
        $csvContent += $rowData -join ","
    }
    
    $csvContent | Out-File -FilePath $csvPath -Encoding UTF8
    Write-Host "Sample exported successfully!" -ForegroundColor Green
    
    # Close Excel
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "`nAnalysis complete. Check $csvPath for data structure." -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}