-- COMPLETE PROPERTIES: Components + PartData + XMLFiles
-- This joins all three tables to get ALL properties including MaxX, MaxY, RawMaterialNumber, etc.

-- Solution 1: Complete properties with highest quantity only
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
    'COMPLETE PROPERTIES - Make Items' as QueryType,
    
    -- Component Properties
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material as ComponentMaterial,
    c.Thickness as ComponentThickness,
    c.AssemblyLevel,
    c.ParentPartNumber,
    
    -- PartData Properties (the ones you're looking for)
    pd.Description as PartDescription,
    pd.MakeBuy as PartMakeBuy,
    pd.Material as PartMaterial,
    pd.Thickness as PartThickness,
    pd.Weight,
    pd.MaxX,
    pd.MaxY, 
    pd.MaxZ,
    pd.Rotation,
    pd.GangQty,
    pd.RawMaterialNumber,
    
    -- XML Source Properties
    xf.PartNumber as XMLPartNumber,
    xf.Release,
    xf.Revision

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
JOIN MaxQuantities mq ON c.ComponentPartNumber = mq.ComponentPartNumber 
    AND c.TotalQuantity = mq.MaxTotalQuantity
LEFT JOIN PartData pd ON pd.PartNumber = c.ComponentPartNumber 
    AND pd.XMLFileID = c.XMLFileID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;

-- Solution 2: Check what PartData exists for our components
SELECT 
    'PartData Availability Check' as QueryType,
    c.ComponentPartNumber,
    COUNT(DISTINCT pd.ID) as PartDataRecords,
    MAX(pd.MaxX) as MaxX_Sample,
    MAX(pd.MaxY) as MaxY_Sample,
    MAX(pd.RawMaterialNumber) as RawMatNum_Sample
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
LEFT JOIN PartData pd ON pd.PartNumber = c.ComponentPartNumber
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
ORDER BY c.ComponentPartNumber;

-- Solution 3: Simple join to see all available columns
SELECT 
    'All Available Columns Sample' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.TotalQuantity,
    pd.MaxX,
    pd.MaxY,
    pd.MaxZ,
    pd.Weight,
    pd.RawMaterialNumber,
    pd.GangQty,
    pd.Rotation,
    xf.Release
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
LEFT JOIN PartData pd ON pd.PartNumber = c.ComponentPartNumber 
    AND pd.XMLFileID = c.XMLFileID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber, xf.Release DESC
LIMIT 20;