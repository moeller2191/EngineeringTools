$excelPath = 'c:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls'
if (Test-Path $excelPath) {
    Write-Host "File exists: $excelPath"
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($excelPath)
        Write-Host "Worksheets in file:"
        for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            $sheetName = $workbook.Worksheets.Item($i).Name
            Write-Host "  $i. $sheetName"
        }
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Host "Error opening Excel file: $($_.Exception.Message)"
    }
} else {
    Write-Host "File does not exist: $excelPath"
}