-- FIXED: ONLY HIGHEST REL - Using actual Release column
-- This uses the Release column directly instead of parsing filenames

-- Step 1: Check what Release values exist (using actual Release column)
SELECT 
    'Available Releases' as Info,
    xf.Release,
    COUNT(*) as ComponentCount
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
GROUP BY xf.Release
ORDER BY xf.Release DESC;

-- Step 2: Get ONLY the highest Release value components
WITH HighestRelease AS (
    -- Find the highest Release value for SULL-1006-0628
    SELECT MAX(xf.Release) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
)
SELECT 
    'ONLY HIGHEST RELEASE Make Items' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.AssemblyLevel,
    c.ParentPartNumber,
    xf.PartNumber as XMLPartNumber,
    xf.Release,
    xf.Revision,
    xf.FileModifiedDate

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;

-- Alternative: Manual approach - specify the release directly
-- Replace 'REL51' with whatever the highest release is from Step 1
SELECT 
    'MANUAL: Specific Release Make Items' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.AssemblyLevel,
    c.ParentPartNumber,
    xf.PartNumber as XMLPartNumber,
    xf.Release

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
    AND xf.Release = 'REL51'  -- Change this to the highest release from Step 1
ORDER BY c.ComponentPartNumber;