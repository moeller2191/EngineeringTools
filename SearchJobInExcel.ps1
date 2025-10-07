# Search for specific job number in Excel data
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open("C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls")
    $worksheet = $workbook.Worksheets["Priority List"]
    
    $searchJob = "IK3NC-0000"
    $totalRows = $worksheet.UsedRange.Rows.Count
    $foundRows = @()
    
    Write-Host "=== Searching for job: $searchJob ===" -ForegroundColor Green
    Write-Host "Total rows to search: $totalRows" -ForegroundColor Yellow
    Write-Host ""
    
    # Search column 3 (fjobno) for the job number
    for ($i = 2; $i -le $totalRows; $i++) {
        $jobCell = $worksheet.Cells.Item($i, 3)
        $jobValue = if ($jobCell.Value2 -ne $null) { $jobCell.Value2.ToString() } else { "" }
        
        if ($jobValue -eq $searchJob) {
            $foundRows += $i
            Write-Host "Found at row $i" -ForegroundColor Green
            
            # Show the full row data for first 10 columns
            for ($col = 1; $col -le 10; $col++) {
                $cell = $worksheet.Cells.Item($i, $col)
                $value = if ($cell.Value2 -ne $null) { $cell.Value2.ToString() } else { "NULL" }
                Write-Host "  Col $col`: '$value'" -ForegroundColor White
            }
            Write-Host ""
        }
        
        # Show progress every 10000 rows
        if ($i % 10000 -eq 0) {
            Write-Host "Searched $i/$totalRows rows..." -ForegroundColor Cyan
        }
    }
    
    if ($foundRows.Count -eq 0) {
        Write-Host "Job '$searchJob' not found in Excel data" -ForegroundColor Red
        Write-Host ""
        Write-Host "Searching for similar jobs (containing 'IK3NC')..." -ForegroundColor Yellow
        
        for ($i = 2; $i -le [Math]::Min(1000, $totalRows); $i++) {
            $jobCell = $worksheet.Cells.Item($i, 3)
            $jobValue = if ($jobCell.Value2 -ne $null) { $jobCell.Value2.ToString() } else { "" }
            
            if ($jobValue -like "*IK3NC*") {
                Write-Host "Similar job found at row $i`: '$jobValue'" -ForegroundColor Cyan
            }
        }
    } else {
        Write-Host "Total occurrences found: $($foundRows.Count)" -ForegroundColor Green
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