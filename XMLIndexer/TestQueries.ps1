# Test Queries for XML Intelligence Database
# Run specific queries to validate the enhanced component extraction logic

Write-Host "=== XML INTELLIGENCE DATABASE TEST QUERIES ===" -ForegroundColor Green
Write-Host ""

# Set the database path
$dbPath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"

# Function to run SQLite queries
function Run-SQLiteQuery {
    param(
        [string]$Query,
        [string]$Description
    )
    
    Write-Host "=== $Description ===" -ForegroundColor Yellow
    Write-Host "Query: $Query" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Use SQLite command line if available, otherwise suggest manual query
        $result = sqlite3.exe $dbPath $Query
        if ($result) {
            $result | ForEach-Object { Write-Host $_ }
        } else {
            Write-Host "No results returned" -ForegroundColor Gray
        }
    } catch {
        Write-Host "Please run this query manually in DB Browser:" -ForegroundColor Red
        Write-Host $Query -ForegroundColor White
    }
    Write-Host ""
    Write-Host "----------------------------------------"
    Write-Host ""
}

# Test 1: Component count
Run-SQLiteQuery -Query "SELECT COUNT(*) as TotalComponents FROM Components;" -Description "Total Components Count"

# Test 2: SULL-1006-0628 Make items (main test case)
Run-SQLiteQuery -Query "SELECT DISTINCT PartNumber, AssemblyName, PDMsmparttoggle, TotalQuantity FROM Components WHERE AssemblyName LIKE '%SULL-1006-0628%' AND PDMsmparttoggle = 'Make' ORDER BY PartNumber;" -Description "SULL-1006-0628 Make Items"

# Test 3: Assembly level breakdown for SULL-1006-0628
Run-SQLiteQuery -Query "SELECT AssemblyLevel, COUNT(*) as Count FROM Components WHERE AssemblyName LIKE '%SULL-1006-0628%' GROUP BY AssemblyLevel ORDER BY AssemblyLevel;" -Description "SULL-1006-0628 Assembly Level Breakdown"

# Test 4: Make/Buy/Stock distribution
Run-SQLiteQuery -Query "SELECT PDMsmparttoggle, COUNT(*) as Count FROM Components GROUP BY PDMsmparttoggle ORDER BY Count DESC;" -Description "Make/Buy/Stock Distribution"

# Test 5: Sample of Stock assemblies to verify they're skipped
Run-SQLiteQuery -Query "SELECT DISTINCT AssemblyName FROM Components WHERE PDMsmparttoggle = 'Stock' LIMIT 10;" -Description "Sample Stock Assemblies"

Write-Host "=== MANUAL QUERIES FOR DB BROWSER ===" -ForegroundColor Green
Write-Host ""
Write-Host "If SQLite command line is not available, please run these queries manually in DB Browser:"
Write-Host ""
Write-Host "1. Component Count:"
Write-Host "   SELECT COUNT(*) as TotalComponents FROM Components;" -ForegroundColor White
Write-Host ""
Write-Host "2. SULL-1006-0628 Make Items:"
Write-Host "   SELECT DISTINCT PartNumber, AssemblyName, PDMsmparttoggle, TotalQuantity" -ForegroundColor White
Write-Host "   FROM Components" -ForegroundColor White
Write-Host "   WHERE AssemblyName LIKE '%SULL-1006-0628%' AND PDMsmparttoggle = 'Make'" -ForegroundColor White
Write-Host "   ORDER BY PartNumber;" -ForegroundColor White
Write-Host ""
Write-Host "3. Assembly Level Breakdown:"
Write-Host "   SELECT AssemblyLevel, COUNT(*) as Count" -ForegroundColor White
Write-Host "   FROM Components" -ForegroundColor White
Write-Host "   WHERE AssemblyName LIKE '%SULL-1006-0628%'" -ForegroundColor White
Write-Host "   GROUP BY AssemblyLevel ORDER BY AssemblyLevel;" -ForegroundColor White