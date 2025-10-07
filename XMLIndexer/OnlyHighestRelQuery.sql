-- ONLY HIGHEST REL - Make Items Properties
-- This query shows ONLY components from the single highest REL release

-- Step 1: Find what REL releases exist for SULL-1006-0628
SELECT 
    'Available REL Releases' as Info,
    xf.Release,
    CASE 
        WHEN xf.PartNumber LIKE '%_REL%' THEN 
            CAST(SUBSTR(xf.PartNumber, INSTR(xf.PartNumber, '_REL') + 4) as INTEGER)
        ELSE 0
    END as ReleaseNumber,
    COUNT(*) as ComponentCount
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
GROUP BY xf.Release, ReleaseNumber
ORDER BY ReleaseNumber DESC;

-- Step 2: Get ONLY the highest REL release components
WITH HighestRel AS (
    -- Find the highest REL number for SULL-1006-0628
    SELECT MAX(
        CASE 
            WHEN xf.PartNumber LIKE '%_REL%' THEN 
                CAST(SUBSTR(xf.PartNumber, INSTR(xf.PartNumber, '_REL') + 4) as INTEGER)
            ELSE 0
        END
    ) as MaxRelease
    FROM Components c
    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
    WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
        AND c.ComponentType = 'Make'
)
SELECT 
    'ONLY HIGHEST REL Make Items' as QueryType,
    -- All Component Properties
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.AssemblyLevel,
    c.ParentPartNumber,
    
    -- XML Properties
    xf.PartNumber as XMLPartNumber,
    xf.Release,
    CASE 
        WHEN xf.PartNumber LIKE '%_REL%' THEN 
            CAST(SUBSTR(xf.PartNumber, INSTR(xf.PartNumber, '_REL') + 4) as INTEGER)
        ELSE 0
    END as ReleaseNumber,
    xf.FileModifiedDate

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
CROSS JOIN HighestRel hr
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
    AND (
        CASE 
            WHEN xf.PartNumber LIKE '%_REL%' THEN 
                CAST(SUBSTR(xf.PartNumber, INSTR(xf.PartNumber, '_REL') + 4) as INTEGER)
            ELSE 0
        END
    ) = hr.MaxRelease
ORDER BY c.ComponentPartNumber;