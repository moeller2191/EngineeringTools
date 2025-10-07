# XML Index Update Script
# Run this periodically to keep your XML database synchronized

param(
    [switch]$Full,
    [switch]$Incremental,
    [switch]$Scheduled
)

# Change to XMLIndexer directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location "$scriptDir\XMLIndexer"

Write-Host "=== XML Database Sync ===" -ForegroundColor Green
Write-Host "Time: $(Get-Date)" -ForegroundColor Gray

if ($Full) {
    Write-Host "Running FULL scan (will reprocess all files)..." -ForegroundColor Yellow
    dotnet run -- --full
}
elseif ($Incremental) {
    Write-Host "Running INCREMENTAL update (new/modified files only)..." -ForegroundColor Cyan
    dotnet run -- --incremental
}
elseif ($Scheduled) {
    Write-Host "Running SCHEDULED update (incremental, silent)..." -ForegroundColor Green
    dotnet run -- --incremental | Out-Host
    Write-Host "Update completed at $(Get-Date)" -ForegroundColor Green
}
else {
    Write-Host "Running SMART update (default mode)..." -ForegroundColor White
    dotnet run
}

Write-Host "`nXML database update completed!" -ForegroundColor Green