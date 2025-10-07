# Debug Excel column headers
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open("C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls")
    $worksheet = $workbook.Worksheets["Priority List"]
    $usedRange = $worksheet.UsedRange
    
    Write-Host "=== Excel Column Headers ===" -ForegroundColor Green
    Write-Host "Total columns: $($usedRange.Columns.Count)" -ForegroundColor Yellow
    Write-Host "Total rows: $($usedRange.Rows.Count)" -ForegroundColor Yellow
    Write-Host ""
    
    # Show first 20 columns
    for ($i = 1; $i -le [Math]::Min(20, $usedRange.Columns.Count); $i++) {
        $headerValue = $worksheet.Cells.Item(1, $i).Value
        Write-Host "Column $i`: '$headerValue'" -ForegroundColor Cyan
    }
    
    Write-Host "" 
    Write-Host "=== Sample Data Row 2 ===" -ForegroundColor Green
    for ($i = 1; $i -le [Math]::Min(10, $usedRange.Columns.Count); $i++) {
        $cellValue = $worksheet.Cells.Item(2, $i).Value
        Write-Host "Col $i`: '$cellValue'" -ForegroundColor White
    }
    
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}