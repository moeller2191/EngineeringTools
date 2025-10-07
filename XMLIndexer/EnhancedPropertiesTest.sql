-- ENHANCED COMPONENT PROPERTIES TEST
-- Test the newly added MaxX, MaxY, RawMaterialNumber, etc. for SULL-1006-0628 components

-- Test 1: Check if component properties are now populated (should show no more NULLs!)
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
    'ENHANCED: Complete Component Properties' as QueryType,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    
    -- NEW ENHANCED PROPERTIES (should no longer be NULL!)
    c.MaxX,
    c.MaxY,
    c.MaxZ,
    c.Weight,
    c.RawMaterialNumber,
    c.GangQty,
    c.Rotation,
    
    xf.Release

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN HighestRelease hr ON xf.Release = hr.MaxRelease
JOIN MaxQuantities mq ON c.ComponentPartNumber = mq.ComponentPartNumber 
    AND c.TotalQuantity = mq.MaxTotalQuantity
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY c.ComponentPartNumber;

-- Test 2: Quick check - count how many components now have properties vs NULLs
SELECT 
    'Property Population Status' as CheckType,
    COUNT(*) as TotalMakeComponents,
    COUNT(c.MaxX) as MaxX_Populated,
    COUNT(c.MaxY) as MaxY_Populated,
    COUNT(c.RawMaterialNumber) as RawMatNum_Populated,
    COUNT(c.Weight) as Weight_Populated
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
    );

-- Test 3: Show just L150237 specifically (the one we were testing before)
SELECT 
    'L150237 Enhanced Properties' as CheckType,
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
    c.Rotation,
    xf.Release
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ComponentPartNumber = 'L150237.SLDPRT'
    AND c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
ORDER BY xf.Release DESC;