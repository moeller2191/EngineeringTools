-- DIAGNOSIS: Why are we seeing duplicates?
-- Let's understand what's causing the duplication

-- Check 1: How many XMLFiles have the same Release for SULL-1006-0628?
SELECT 
    'XMLFiles per Release' as CheckType,
    xf.Release,
    COUNT(DISTINCT xf.ID) as XMLFileCount,
    COUNT(DISTINCT xf.PartNumber) as UniquePartNumbers
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
GROUP BY xf.Release
ORDER BY xf.Release DESC;

-- Check 2: Show the actual XMLFiles for the highest release
WITH HighestRelease AS (
    SELECT MAX(xf.Release) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
)
SELECT 
    'XMLFiles in Highest Release' as CheckType,
    DISTINCT xf.ID as XMLFileID,
    xf.PartNumber as XMLPartNumber,
    xf.Release,
    xf.Revision,
    COUNT(*) as ComponentsInThisXML
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
GROUP BY xf.ID, xf.PartNumber, xf.Release, xf.Revision
ORDER BY xf.PartNumber;

-- Check 3: Show duplicate components across XMLFiles
SELECT 
    'Duplicate Components Analysis' as CheckType,
    c.ComponentPartNumber,
    COUNT(DISTINCT c.XMLFileID) as AppearsInXMLFiles,
    GROUP_CONCAT(DISTINCT xf.PartNumber) as XMLPartNumbers
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
    AND xf.Release = (
        SELECT MAX(xf2.Release)
        FROM Components c2
        JOIN XMLFiles xf2 ON c2.XMLFileID = xf2.ID
        WHERE c2.ParentPartNumber LIKE '%SULL-1006-0628%' 
            AND c2.ComponentType = 'Make'
    )
GROUP BY c.ComponentPartNumber
HAVING COUNT(DISTINCT c.XMLFileID) > 1
ORDER BY c.ComponentPartNumber;

-- FIXED SOLUTION: True DISTINCT components from highest release
-- This removes ALL duplicates by using DISTINCT on component properties
WITH HighestRelease AS (
    SELECT MAX(xf.Release) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
)
SELECT DISTINCT
    'UNIQUE Make Items - No Duplicates' as QueryType,
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
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;