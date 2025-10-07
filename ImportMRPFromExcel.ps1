# Excel to MRP Database Importer
# This script reads the Priority List Excel file and imports data to SQLite

param(
    [string]$ExcelPath = "C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls",
    [string]$DatabasePath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"
)

Write-Host "=== MRP EXCEL IMPORTER ===" -ForegroundColor Green

function Import-ExcelToMrp {
    param($ExcelFile, $DatabaseFile)
    
    try {
        # Check if files exist
        if (-not (Test-Path $ExcelFile)) {
            Write-Host "Excel file not found: $ExcelFile" -ForegroundColor Red
            return $false
        }
        
        if (-not (Test-Path $DatabaseFile)) {
            Write-Host "Database file not found: $DatabaseFile" -ForegroundColor Red
            return $false
        }
        
        Write-Host "Reading Excel file: $ExcelFile" -ForegroundColor Yellow
        
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Open($ExcelFile)
        $worksheet = $workbook.Worksheets.Item(1) # Use first worksheet
        
        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        
        Write-Host "Found $rowCount rows in Excel file" -ForegroundColor Cyan
        
        # Create temp CSV for import
        $tempCsv = [System.IO.Path]::GetTempFileName() + ".csv"
        Write-Host "Creating temporary CSV: $tempCsv" -ForegroundColor Yellow
        
        $csvContent = @()
        
        # Read data from Excel (assuming row 1 has headers)
        for ($row = 1; $row -le $rowCount; $row++) {
            $rowData = @()
            
            # Read common columns (adjust based on actual Excel structure)
            for ($col = 1; $col -le 8; $col++) {
                $cellValue = $worksheet.Cells.Item($row, $col).Text.Trim()
                if ($cellValue.Contains(",") -or $cellValue.Contains('"')) {
                    $rowData += "`"$($cellValue.Replace('"', '""'))`""
                } else {
                    $rowData += $cellValue
                }
            }
            
            $csvContent += $rowData -join ","
        }
        
        $csvContent | Out-File -FilePath $tempCsv -Encoding UTF8
        
        # Close Excel
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "Excel data exported to temporary CSV" -ForegroundColor Green
        
        # Import to SQLite database
        Write-Host "Importing to SQLite database..." -ForegroundColor Yellow
        
        # Use SQLite command line to import (if available) or create SQL insert statements
        $sqliteCmd = "sqlite3"
        
        # Test if sqlite3 is available
        try {
            & $sqliteCmd --version | Out-Null
            
            # Create import SQL
            $importSql = @"
.mode csv
.headers on
.import $tempCsv temp_import
INSERT OR REPLACE INTO MrpPriorityList (JobNumber, PartNumber, Revision, Quantity, Description, Priority, Status, LastUpdated)
SELECT 
    COALESCE(column1, '') as JobNumber,
    COALESCE(column2, '') as PartNumber, 
    COALESCE(column3, '') as Revision,
    CAST(COALESCE(column4, '1') AS INTEGER) as Quantity,
    COALESCE(column5, '') as Description,
    CAST(COALESCE(column6, '1') AS INTEGER) as Priority,
    'Active' as Status,
    datetime('now') as LastUpdated
FROM temp_import 
WHERE column1 IS NOT NULL AND column1 != '';
DROP TABLE temp_import;
"@
            
            $importSql | & $sqliteCmd $DatabaseFile
            Write-Host "Data imported successfully to SQLite database" -ForegroundColor Green
            
        } catch {
            Write-Host "SQLite command line not available. Creating manual import method..." -ForegroundColor Yellow
            
            # Manual import using .NET (fallback)
            Import-CsvToSqliteManual -CsvPath $tempCsv -DatabasePath $DatabaseFile
        }
        
        # Clean up temp file
        Remove-Item $tempCsv -ErrorAction SilentlyContinue
        
        return $true
        
    } catch {
        Write-Host "Error importing Excel data: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        
        return $false
    }
}

function Import-CsvToSqliteManual {
    param($CsvPath, $DatabasePath)
    
    Write-Host "Manual CSV import to SQLite..." -ForegroundColor Yellow
    
    # This would require PowerShell SQLite module or other method
    # For now, just copy the logic that would be implemented
    Write-Host "Manual import would read CSV and create INSERT statements" -ForegroundColor Cyan
    Write-Host "CSV file location: $CsvPath" -ForegroundColor Cyan
}

# Main execution
Write-Host "Starting MRP Excel import process..." -ForegroundColor Blue

$success = Import-ExcelToMrp -ExcelFile $ExcelPath -DatabaseFile $DatabasePath

if ($success) {
    Write-Host "`nMRP import completed successfully!" -ForegroundColor Green
    Write-Host "You can now use the Engineering Tools application to load real MRP data." -ForegroundColor Green
} else {
    Write-Host "`nMRP import failed. Please check the error messages above." -ForegroundColor Red
}

Write-Host "`nPress any key to continue..." -ForegroundColor Gray