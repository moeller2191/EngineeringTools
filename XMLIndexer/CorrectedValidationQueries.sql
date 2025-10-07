-- CORRECTED Enhanced MRP Logic Validation Queries
-- Run these in DB Browser for SQLite to verify the new assembly hierarchy logic

-- 1. Total component count (should be significantly less than before)
SELECT 'Total Components in Database:' as Description, COUNT(*) as Count FROM Components;

-- 2. SULL-1006-0628 Make Items (CORRECTED QUERY - should NOT include items from stock sub-assemblies)
SELECT 
    'SULL-1006-0628 Make Items' as QueryType,
    ComponentPartNumber, 
    ParentPartNumber as AssemblyName, 
    ComponentType as PDMsmparttoggle, 
    TotalQuantity,
    AssemblyLevel
FROM Components 
WHERE ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND ComponentType = 'Make' 
ORDER BY ComponentPartNumber;

-- 3. Alternative: Search both parent and component for SULL-1006-0628
SELECT 
    'SULL-1006-0628 All Related Items' as QueryType,
    ComponentPartNumber, 
    ParentPartNumber as AssemblyName, 
    ComponentType, 
    TotalQuantity,
    AssemblyLevel
FROM Components 
WHERE (ParentPartNumber LIKE '%SULL-1006-0628%' OR ComponentPartNumber LIKE '%SULL-1006-0628%')
    AND ComponentType = 'Make' 
ORDER BY ComponentPartNumber;

-- 4. Compare: Before vs After component extraction
SELECT 
    'Component Extraction Comparison' as QueryType,
    ComponentType as Type, 
    COUNT(*) as Count 
FROM Components 
GROUP BY ComponentType 
ORDER BY Count DESC;

-- 5. Verify no components from Stock assemblies leak through
SELECT 
    'Stock Assembly Verification' as QueryType,
    ParentPartNumber as AssemblyName,
    COUNT(*) as ComponentCount
FROM Components 
WHERE ParentPartNumber IN (
    SELECT ParentPartNumber 
    FROM Components 
    WHERE ComponentType = 'Stock'
) 
AND ComponentType != 'Stock'
GROUP BY ParentPartNumber
HAVING ComponentCount > 0
LIMIT 10;

-- 6. Sample of successfully extracted Make items
SELECT 
    'Sample Make Items' as QueryType,
    ComponentPartNumber, 
    ParentPartNumber as AssemblyName, 
    ComponentType, 
    TotalQuantity
FROM Components 
WHERE ComponentType = 'Make' 
LIMIT 10;

-- 7. Check what SULL-1006-0628 items exist in the database
SELECT 
    'All SULL-1006-0628 References' as QueryType,
    ComponentPartNumber, 
    ParentPartNumber, 
    ComponentType, 
    TotalQuantity
FROM Components 
WHERE (ParentPartNumber LIKE '%SULL-1006-0628%' OR ComponentPartNumber LIKE '%SULL-1006-0628%')
ORDER BY ComponentType, ComponentPartNumber;