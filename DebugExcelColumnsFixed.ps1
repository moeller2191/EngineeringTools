# Fixed Excel column headers debug script
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open("C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls")
    $worksheet = $workbook.Worksheets["Priority List"]
    
    Write-Host "=== Excel Column Headers ===" -ForegroundColor Green
    Write-Host "Total columns: $($worksheet.UsedRange.Columns.Count)" -ForegroundColor Yellow
    Write-Host "Total rows: $($worksheet.UsedRange.Rows.Count)" -ForegroundColor Yellow
    Write-Host ""
    
    # Show first 30 columns with proper value extraction
    for ($i = 1; $i -le 30; $i++) {
        $headerCell = $worksheet.Cells.Item(1, $i)
        $headerValue = if ($headerCell.Value2 -ne $null) { $headerCell.Value2.ToString() } else { "NULL" }
        Write-Host "Column $i`: '$headerValue'" -ForegroundColor Cyan
    }
    
    Write-Host "" 
    Write-Host "=== Sample Data Row 2 ===" -ForegroundColor Green
    for ($i = 1; $i -le 30; $i++) {
        $dataCell = $worksheet.Cells.Item(2, $i)
        $cellValue = if ($dataCell.Value2 -ne $null) { $dataCell.Value2.ToString() } else { "NULL" }
        Write-Host "Col $i`: '$cellValue'" -ForegroundColor White
    }
    
    Write-Host "" 
    Write-Host "=== Sample Data Row 3 ===" -ForegroundColor Green
    for ($i = 1; $i -le 30; $i++) {
        $dataCell = $worksheet.Cells.Item(3, $i)
        $cellValue = if ($dataCell.Value2 -ne $null) { $dataCell.Value2.ToString() } else { "NULL" }
        Write-Host "Col $i`: '$cellValue'" -ForegroundColor White
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}