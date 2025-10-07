# Check revision data for specific job IK3NC-0000
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open("C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls")
    $worksheet = $workbook.Worksheets["Priority List"]
    
    $searchJob = "IK3NC-0000"
    $totalRows = $worksheet.UsedRange.Rows.Count
    
    Write-Host "=== Checking revision data for job: $searchJob ===" -ForegroundColor Green
    Write-Host ""
    
    # Search column 3 (fjobno) for the job number and show revision data
    for ($i = 2; $i -le $totalRows; $i++) {
        $jobCell = $worksheet.Cells.Item($i, 3)
        $jobValue = if ($jobCell.Value2 -ne $null) { $jobCell.Value2.ToString() } else { "" }
        
        if ($jobValue -eq $searchJob) {
            Write-Host "Found at row $i" -ForegroundColor Green
            
            # Show key columns including revision
            $partNum = if ($worksheet.Cells.Item($i, 5).Value2 -ne $null) { $worksheet.Cells.Item($i, 5).Value2.ToString() } else { "NULL" }
            $desc = if ($worksheet.Cells.Item($i, 6).Value2 -ne $null) { $worksheet.Cells.Item($i, 6).Value2.ToString() } else { "NULL" }
            $revision = if ($worksheet.Cells.Item($i, 23).Value2 -ne $null) { $worksheet.Cells.Item($i, 23).Value2.ToString() } else { "NULL" }
            $memo = if ($worksheet.Cells.Item($i, 22).Value2 -ne $null) { $worksheet.Cells.Item($i, 22).Value2.ToString() } else { "NULL" }
            
            Write-Host "  Job: $jobValue" -ForegroundColor White
            Write-Host "  Part: $partNum" -ForegroundColor White
            Write-Host "  Description: $desc" -ForegroundColor White
            Write-Host "  Revision (Col 23): $revision" -ForegroundColor Yellow
            Write-Host "  Memo (Col 22): $memo" -ForegroundColor Cyan
            Write-Host ""
            
            break  # Just show first occurrence
        }
        
        # Show progress every 10000 rows
        if ($i % 10000 -eq 0) {
            Write-Host "Searched $i/$totalRows rows..." -ForegroundColor Cyan
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}