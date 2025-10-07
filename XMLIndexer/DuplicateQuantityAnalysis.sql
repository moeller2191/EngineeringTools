-- INVESTIGATE: Why L155471 appears multiple times
-- Let's see where this part appears and why it has different quantities

-- Check 1: Show all instances of L155471 across XMLFiles
SELECT 
    'L155471 Analysis' as CheckType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.Quantity,
    c.TotalQuantity,
    c.AssemblyLevel,
    c.ParentPartNumber,
    xf.PartNumber as XMLPartNumber,
    xf.Release,
    c.XMLFileID
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ComponentPartNumber = 'L155471.SLDPRT'
    AND c.ComponentType = 'Make'
    AND c.ParentPartNumber LIKE '%SULL-1006-0628%'
ORDER BY xf.Release DESC, c.Quantity DESC;

-- Check 2: Show all duplicate components (same part, different quantities)
WITH HighestRelease AS (
    SELECT MAX(xf.Release) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
)
SELECT 
    'Duplicate Quantities Analysis' as CheckType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    COUNT(*) as Occurrences,
    GROUP_CONCAT(DISTINCT c.Quantity) as DifferentQuantities,
    GROUP_CONCAT(DISTINCT c.TotalQuantity) as DifferentTotalQuantities,
    GROUP_CONCAT(DISTINCT xf.PartNumber) as XMLFiles
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
GROUP BY c.ComponentPartNumber, c.ComponentDescription
HAVING COUNT(*) > 1
ORDER BY c.ComponentPartNumber;

-- SOLUTION 1: Aggregate to get TOTAL quantities (sum all instances)
WITH HighestRelease AS (
    SELECT MAX(xf.Release) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
)
SELECT 
    'AGGREGATED Make Items (Total Quantities)' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    SUM(c.Quantity) as TotalQuantity_Sum,
    MAX(c.TotalQuantity) as TotalQuantity_Max,
    c.Material,
    c.Thickness,
    MIN(c.AssemblyLevel) as MinAssemblyLevel,
    c.ParentPartNumber
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
GROUP BY c.ComponentPartNumber, c.ComponentDescription, c.ComponentType, c.Material, c.Thickness, c.ParentPartNumber
ORDER BY c.ComponentPartNumber;

-- SOLUTION 2: Keep only the highest quantity instance
WITH HighestRelease AS (
    SELECT MAX(xf.Release) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
),
MaxQuantities AS (
    SELECT 
        c.ComponentPartNumber,
        MAX(c.TotalQuantity) as MaxTotalQuantity
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
    GROUP BY c.ComponentPartNumber
)
SELECT DISTINCT
    'HIGHEST QUANTITY Make Items Only' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.AssemblyLevel,
    c.ParentPartNumber
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
JOIN MaxQuantities mq ON c.ComponentPartNumber = mq.ComponentPartNumber 
    AND c.TotalQuantity = mq.MaxTotalQuantity
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;