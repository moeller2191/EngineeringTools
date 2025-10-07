# Quick test script to verify .NET installation and XML access
# Run this after installing .NET to make sure everything is ready

Write-Host "=== .NET and XML Access Test ===" -ForegroundColor Green
Write-Host

# Test 1: Check .NET installation
Write-Host "1. Testing .NET installation..." -ForegroundColor Yellow
try {
    $dotnetVersion = dotnet --version
    Write-Host "   ✓ .NET SDK installed: $dotnetVersion" -ForegroundColor Green
} catch {
    Write-Host "   ✗ .NET SDK not found or not in PATH" -ForegroundColor Red
    Write-Host "   Make sure to restart PowerShell after installing .NET" -ForegroundColor Yellow
    exit 1
}

# Test 2: Check XML folder access
Write-Host
Write-Host "2. Testing XML folder access..." -ForegroundColor Yellow

$xmlPaths = @(
    "\\kmi-solidworks22\solidworks22common\CUT LIST XML",
    "\\kmi-solidworks22\solidworks22common\CUT LIST XML\Legacy", 
    "\\kmi-solidworks22\solidworks22common\CUT LIST XML\New"
)

$totalXmlFiles = 0
foreach ($path in $xmlPaths) {
    Write-Host "   Checking: $path" -ForegroundColor Cyan
    if (Test-Path $path) {
        $xmlFiles = Get-ChildItem -Path $path -Filter "*.xml" -ErrorAction SilentlyContinue
        $count = $xmlFiles.Count
        $totalXmlFiles += $count
        Write-Host "     ✓ Accessible - Found $count XML files" -ForegroundColor Green
        
        # Show a sample filename if any exist
        if ($count -gt 0) {
            $sample = $xmlFiles[0].Name
            Write-Host "     Sample: $sample" -ForegroundColor Gray
        }
    } else {
        Write-Host "     ✗ Not accessible or doesn't exist" -ForegroundColor Red
    }
}

Write-Host
Write-Host "=== Summary ===" -ForegroundColor Green
Write-Host "Total XML files found: $totalXmlFiles" -ForegroundColor White

if ($totalXmlFiles -gt 0) {
    Write-Host "✓ Ready to process XML files!" -ForegroundColor Green
    Write-Host
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. cd XMLIndexer" -ForegroundColor White
    Write-Host "2. dotnet build" -ForegroundColor White
    Write-Host "3. dotnet run" -ForegroundColor White
} else {
    Write-Host "⚠ No XML files found. Check network access and paths." -ForegroundColor Yellow
}

Write-Host
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")