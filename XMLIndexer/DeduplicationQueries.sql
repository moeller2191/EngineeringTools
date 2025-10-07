-- Understanding Multiple XML References for SULL-1006-0628
-- These queries will help you see why the same make item appears multiple times

-- 1. How many XML files contain SULL-1006-0628?
SELECT 
    'XML Files containing SULL-1006-0628' as QueryType,
    COUNT(DISTINCT XMLFileID) as UniqueXMLFiles,
    COUNT(*) as TotalReferences
FROM Components 
WHERE (ParentPartNumber LIKE '%SULL-1006-0628%' OR ComponentPartNumber LIKE '%SULL-1006-0628%');

-- 2. Which XML files contain SULL-1006-0628?
SELECT 
    'XML Files with SULL-1006-0628' as QueryType,
    xf.FilePath,
    xf.PartNumber as XMLPartNumber,
    COUNT(*) as ComponentCount
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE (c.ParentPartNumber LIKE '%SULL-1006-0628%' OR c.ComponentPartNumber LIKE '%SULL-1006-0628%')
GROUP BY xf.ID, xf.FilePath, xf.PartNumber
ORDER BY ComponentCount DESC;

-- 3. DEDUPLICATED Make items for SULL-1006-0628 (removes duplicates)
SELECT DISTINCT
    'UNIQUE Make Items for SULL-1006-0628' as QueryType,
    ComponentPartNumber, 
    ComponentType,
    MAX(TotalQuantity) as MaxQuantity,
    COUNT(DISTINCT XMLFileID) as AppearsInXMLs
FROM Components 
WHERE ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND ComponentType = 'Make' 
GROUP BY ComponentPartNumber, ComponentType
ORDER BY ComponentPartNumber;

-- 4. Show duplicates with their XML sources
SELECT 
    'Make Items with XML Sources' as QueryType,
    c.ComponentPartNumber, 
    c.ComponentType,
    c.TotalQuantity,
    xf.PartNumber as XMLPartNumber,
    xf.FilePath
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber, xf.PartNumber;