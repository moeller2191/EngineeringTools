-- Complete Properties of Make Items from Highest REL Release
-- Shows ALL columns/properties for Make items from SULL-1006-0628 highest release

WITH ParsedReleases AS (
    SELECT 
        ID,
        PartNumber,
        FilePath,
        LastModified,
        -- Extract base part number (everything before _REL)
        CASE 
            WHEN PartNumber LIKE '%_REL%' THEN SUBSTR(PartNumber, 1, INSTR(PartNumber, '_REL') - 1)
            ELSE PartNumber 
        END as BasePartNumber,
        -- Extract REL number (everything after _REL)
        CASE 
            WHEN PartNumber LIKE '%_REL%' THEN 
                CAST(SUBSTR(PartNumber, INSTR(PartNumber, '_REL') + 4) as INTEGER)
            ELSE 0
        END as ReleaseNumber
    FROM XMLFiles 
    WHERE PartNumber LIKE '%SULL-1006-0628%'
),
HighestReleases AS (
    SELECT 
        BasePartNumber,
        MAX(ReleaseNumber) as HighestRelease
    FROM ParsedReleases
    GROUP BY BasePartNumber
)
SELECT DISTINCT
    -- Component Properties
    c.ID as ComponentID,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.ComponentIndex,
    c.AssemblyLevel,
    
    -- Parent Assembly Properties  
    c.ParentPartNumber,
    
    -- XML File Properties
    c.XMLFileID,
    pr.PartNumber as XMLPartNumber,
    pr.ReleaseNumber,
    pr.FilePath,
    pr.LastModified,
    
    -- Additional XML File Data (joined from XMLFiles table)
    xf.PartNumber as XMLFilePartNumber,
    xf.Material as XMLFileMaterial,
    xf.Thickness as XMLFileThickness,
    xf.FileSize,
    xf.ProcessedDate

FROM Components c
JOIN ParsedReleases pr ON c.XMLFileID = pr.ID
JOIN HighestReleases hr ON pr.BasePartNumber = hr.BasePartNumber 
    AND pr.ReleaseNumber = hr.HighestRelease
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber, c.AssemblyLevel;

-- Alternative: Simpler query with all component properties (no REL filtering for comparison)
SELECT 
    'All Properties - All Make Items' as QueryType,
    c.*,  -- All component columns
    xf.PartNumber as XMLPartNumber,
    xf.FilePath,
    xf.Material as XMLMaterial,
    xf.Thickness as XMLThickness,
    xf.LastModified,
    xf.FileSize
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber
LIMIT 20;  -- Limit to see structure first

-- Schema check: Show all available columns
PRAGMA table_info(Components);
PRAGMA table_info(XMLFiles);