# Force fresh Excel import by clearing database first
$dbPath = "C:\Scripts\EngineeringTools\XMLIndex.db"

Write-Host "=== Forcing Fresh Excel Import ===" -ForegroundColor Green
Write-Host "This will clear the existing database and force a fresh import"
Write-Host ""

if (Test-Path $dbPath) {
    Write-Host "Removing existing database file..." -ForegroundColor Yellow
    Remove-Item $dbPath -Force
    Write-Host "Database file removed" -ForegroundColor Green
} else {
    Write-Host "Database file not found at: $dbPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Now restart the application or click 'Load from Excel' to trigger a fresh import" -ForegroundColor Cyan
Write-Host "The new import should include the revision field from column 23 (fcudrev)" -ForegroundColor Cyan