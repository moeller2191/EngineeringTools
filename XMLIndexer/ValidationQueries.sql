-- Enhanced MRP Logic Validation Queries
-- Run these in DB Browser for SQLite to verify the new assembly hierarchy logic

-- 1. Total component count (should be significantly less than before)
SELECT 'Total Components in Database:' as Description, COUNT(*) as Count FROM Components;

-- 2. SULL-1006-0628 Make Items (main test - should NOT include items from stock sub-assemblies)
SELECT 
    'SULL-1006-0628 Make Items' as QueryType,
    PartNumber, 
    AssemblyName, 
    PDMsmparttoggle, 
    TotalQuantity,
    AssemblyLevel
FROM Components 
WHERE AssemblyName LIKE '%SULL-1006-0628%' 
    AND PDMsmparttoggle = 'Make' 
ORDER BY PartNumber;

-- 3. Compare: Before vs After component extraction
SELECT 
    'Component Extraction Comparison' as QueryType,
    PDMsmparttoggle as Type, 
    COUNT(*) as Count 
FROM Components 
GROUP BY PDMsmparttoggle 
ORDER BY Count DESC;

-- 4. Verify no components from Stock assemblies leak through
SELECT 
    'Stock Assembly Verification' as QueryType,
    AssemblyName,
    COUNT(*) as ComponentCount
FROM Components 
WHERE AssemblyName IN (
    SELECT AssemblyName 
    FROM Components 
    WHERE PDMsmparttoggle = 'Stock'
) 
AND PDMsmparttoggle != 'Stock'
GROUP BY AssemblyName
HAVING ComponentCount > 0
LIMIT 10;

-- 5. Sample of successfully extracted Make items
SELECT 
    'Sample Make Items' as QueryType,
    PartNumber, 
    AssemblyName, 
    PDMsmparttoggle, 
    TotalQuantity
FROM Components 
WHERE PDMsmparttoggle = 'Make' 
LIMIT 10;