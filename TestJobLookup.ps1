# Test script to verify job lookup functionality
param(
    [string]$JobNumber = "IK3NC-0000"
)

$dbPath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"

if (-not (Test-Path $dbPath)) {
    Write-Error "Database not found at: $dbPath"
    exit 1
}

# Test the JOIN query directly
$query = @"
SELECT m.*, 
       COALESCE(x.XmlStatus, 'No XML') as XmlStatus,
       COALESCE(x.HighestRelease, 0) as HighestRelease,
       COALESCE(x.ComponentCount, 0) as ComponentCount
FROM MrpPriorityList m
LEFT JOIN XMLIndex x ON m.PartNumber = x.PartNumber AND m.Revision = x.Revision
WHERE m.JobNumber LIKE '%$JobNumber%'
"@

Write-Host "Testing job lookup for: $JobNumber"
Write-Host "Database path: $dbPath"
Write-Host "Query:" -ForegroundColor Yellow
Write-Host $query

# Use sqlite3 if available, otherwise indicate the query to run
try {
    $result = sqlite3.exe $dbPath $query 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nResults:" -ForegroundColor Green
        $result | ForEach-Object { Write-Host $_ }
    } else {
        Write-Host "SQLite3 not found in PATH. Please run the query manually in a SQLite browser." -ForegroundColor Yellow
    }
} catch {
    Write-Host "SQLite3 not available. Query prepared for manual execution." -ForegroundColor Yellow
}

# Also test the MrpPriorityList table structure
$structureQuery = "PRAGMA table_info(MrpPriorityList);"
Write-Host "`nTable structure:" -ForegroundColor Cyan
try {
    $structure = sqlite3.exe $dbPath $structureQuery 2>$null
    if ($LASTEXITCODE -eq 0) {
        $structure | ForEach-Object { Write-Host $_ }
    }
} catch {
    Write-Host "Run manually: $structureQuery"
}