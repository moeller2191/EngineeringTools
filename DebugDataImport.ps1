# Debug script to check Excel file contents for data import
param(
    [string]$ExcelPath = "c:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"
)

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "Opening Excel file: $ExcelPath" -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    # List all worksheets
    Write-Host "`nWorksheets found:" -ForegroundColor Green
    for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
        $worksheet = $workbook.Worksheets.Item($i)
        Write-Host "  $i. $($worksheet.Name)" -ForegroundColor Cyan
        
        # Check used range
        $usedRange = $worksheet.UsedRange
        if ($usedRange) {
            $lastRow = $usedRange.Rows.Count
            $lastCol = $usedRange.Columns.Count
            Write-Host "     Used Range: $lastRow rows x $lastCol columns" -ForegroundColor Gray
            
            # Show headers (first row)
            Write-Host "     Headers:" -ForegroundColor Gray
            for ($col = 1; $col -le [Math]::Min($lastCol, 15); $col++) {
                $header = $worksheet.Cells.Item(1, $col).Value2
                if ($header) {
                    Write-Host "       Col $col`: $header" -ForegroundColor White
                }
            }
            
            # Show a few sample data rows
            Write-Host "     Sample Data (rows 2-4):" -ForegroundColor Gray
            for ($row = 2; $row -le [Math]::Min($lastRow, 4); $row++) {
                $rowData = @()
                for ($col = 1; $col -le [Math]::Min($lastCol, 5); $col++) {
                    $cellValue = $worksheet.Cells.Item($row, $col).Value2
                    $rowData += if ($cellValue) { $cellValue.ToString() } else { "[empty]" }
                }
                Write-Host "       Row $row`: $($rowData -join ' | ')" -ForegroundColor White
            }
        }
        Write-Host ""
    }
    
    # Close workbook
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($excel) {
        try { $excel.Quit() } catch {}
    }
}