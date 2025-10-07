-- REMOVE DUPLICATES - Show only entries WITH properties populated
-- This eliminates the duplicate rows that have NULL/empty properties

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
),
ComponentsWithProperties AS (
    -- Only get entries that have properties populated (not NULL)
    SELECT DISTINCT
        c.ComponentPartNumber,
        c.ComponentDescription,
        c.TotalQuantity,
        c.Material,
        c.Thickness,
        c.MaxX,
        c.MaxY,
        c.MaxZ,
        c.Weight,
        c.RawMaterialNumber,
        c.GangQty,
        c.Rotation
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
    JOIN MaxQuantities mq ON c.ComponentPartNumber = mq.ComponentPartNumber 
        AND c.TotalQuantity = mq.MaxTotalQuantity
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
        AND c.MaxX IS NOT NULL  -- Only entries with properties
        AND c.RawMaterialNumber IS NOT NULL
)
SELECT 
    'CLEAN: No Duplicates - Only With Properties' as QueryType,
    ComponentPartNumber,
    ComponentDescription,
    TotalQuantity,
    Material,
    Thickness,
    MaxX,
    MaxY,
    MaxZ,
    Weight,
    RawMaterialNumber,
    GangQty,
    Rotation
FROM ComponentsWithProperties
ORDER BY ComponentPartNumber;

-- Alternative: Simple approach - just filter out NULLs
SELECT DISTINCT
    'SIMPLE: Properties Only' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.MaxX,
    c.MaxY,
    c.MaxZ,
    c.Weight,
    c.RawMaterialNumber,
    c.GangQty,
    c.Rotation
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
    AND c.MaxX IS NOT NULL
    AND c.RawMaterialNumber IS NOT NULL
    AND xf.Release = (
        SELECT MAX(xf2.Release)
        FROM Components c2
        JOIN XMLFiles xf2 ON c2.XMLFileID = xf2.ID
        WHERE c2.ParentPartNumber LIKE '%SULL-1006-0628%' 
            AND c2.ComponentType = 'Make'
    )
ORDER BY c.ComponentPartNumber;