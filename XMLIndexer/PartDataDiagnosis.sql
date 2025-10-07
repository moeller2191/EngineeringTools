-- DIAGNOSE: Why PartData is NULL for components
-- Let's understand why the PartData LEFT JOIN is failing

-- Check 1: Do these parts exist in PartData at all?
SELECT 
    'PartData Existence Check' as CheckType,
    c.ComponentPartNumber as ComponentPart,
    pd.PartNumber as PartDataPart,
    COUNT(pd.ID) as PartDataRecords
FROM (
    SELECT DISTINCT ComponentPartNumber 
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
        AND xf.Release = '53'
) c
LEFT JOIN PartData pd ON pd.PartNumber = c.ComponentPartNumber
GROUP BY c.ComponentPartNumber, pd.PartNumber
ORDER BY c.ComponentPartNumber;

-- Check 2: Look for PartData with similar names (without .SLDPRT extension)
SELECT 
    'PartData Name Variations' as CheckType,
    c.ComponentPartNumber,
    pd.PartNumber as FoundInPartData,
    pd.Description,
    pd.MaxX,
    pd.MaxY,
    pd.RawMaterialNumber
FROM (
    SELECT DISTINCT ComponentPartNumber 
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
        AND xf.Release = '53'
    LIMIT 5  -- Just test first 5
) c
LEFT JOIN PartData pd ON pd.PartNumber = REPLACE(c.ComponentPartNumber, '.SLDPRT', '')
WHERE pd.PartNumber IS NOT NULL;

-- Check 3: What PartData actually exists for XMLFileIDs related to SULL-1006-0628?
SELECT 
    'Available PartData for SULL-1006-0628 XMLFiles' as CheckType,
    pd.PartNumber,
    pd.Description,
    pd.MaxX,
    pd.MaxY,
    pd.RawMaterialNumber,
    xf.PartNumber as XMLPartNumber,
    xf.Release
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN PartData pd ON pd.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%'
    AND xf.Release = '53'
ORDER BY pd.PartNumber;

-- SOLUTION 1: Try matching without .SLDPRT extension
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
    'FIXED: Try without .SLDPRT extension' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.TotalQuantity,
    c.Material as ComponentMaterial,
    c.Thickness as ComponentThickness,
    
    -- Try PartData match without extension
    pd.MaxX,
    pd.MaxY, 
    pd.MaxZ,
    pd.Weight,
    pd.RawMaterialNumber,
    pd.GangQty,
    
    xf.Release

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
JOIN MaxQuantities mq ON c.ComponentPartNumber = mq.ComponentPartNumber 
    AND c.TotalQuantity = mq.MaxTotalQuantity
LEFT JOIN PartData pd ON pd.PartNumber = REPLACE(c.ComponentPartNumber, '.SLDPRT', '')
    AND pd.XMLFileID = c.XMLFileID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;

-- SOLUTION 2: Try ANY PartData match for the part (ignore XMLFileID)
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
    'FIXED: Any PartData match' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.TotalQuantity,
    
    -- Get ANY PartData for this part
    pd.MaxX,
    pd.MaxY, 
    pd.MaxZ,
    pd.Weight,
    pd.RawMaterialNumber,
    pd.GangQty,
    
    xf.Release

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
JOIN MaxQuantities mq ON c.ComponentPartNumber = mq.ComponentPartNumber 
    AND c.TotalQuantity = mq.MaxTotalQuantity
LEFT JOIN PartData pd ON (
    pd.PartNumber = c.ComponentPartNumber OR 
    pd.PartNumber = REPLACE(c.ComponentPartNumber, '.SLDPRT', '')
)
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;