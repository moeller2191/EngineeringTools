-- CORRECTED: Complete Properties of Make Items from Highest REL Release
-- Shows ALL columns/properties for Make items from SULL-1006-0628 highest release

WITH ParsedReleases AS (
    SELECT 
        ID,
        PartNumber,
        FilePath,
        FileName,
        Revision,
        Release,
        FileModifiedDate,
        ParsedDate,
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
    -- Component Properties (ALL COLUMNS)
    c.ID as ComponentID,
    c.XMLFileID,
    c.ParentPartNumber,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.ComponentIndex,
    c.AssemblyLevel,
    
    -- XML Source Properties (ALL COLUMNS)
    pr.PartNumber as XMLPartNumber,
    pr.FileName as XMLFileName,
    pr.Revision as XMLRevision,
    pr.Release as XMLRelease,
    pr.ReleaseNumber as ExtractedRelNumber,
    pr.FilePath as XMLFilePath,
    pr.FileModifiedDate,
    pr.ParsedDate

FROM Components c
JOIN ParsedReleases pr ON c.XMLFileID = pr.ID
JOIN HighestReleases hr ON pr.BasePartNumber = hr.BasePartNumber 
    AND pr.ReleaseNumber = hr.HighestRelease
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber, c.AssemblyLevel;

-- Alternative: Simpler version without REL parsing (for testing)
SELECT 
    -- All Component Properties
    c.ID as ComponentID,
    c.XMLFileID,
    c.ParentPartNumber,
    c.ComponentPartNumber,
    c.ComponentDescription,
    c.ComponentType,
    c.Quantity,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    c.ComponentIndex,
    c.AssemblyLevel,
    
    -- All XML Properties
    xf.PartNumber as XMLPartNumber,
    xf.FileName,
    xf.Revision,
    xf.Release,
    xf.FilePath,
    xf.FileModifiedDate,
    xf.ParsedDate

FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber, xf.Release DESC
LIMIT 50;  -- Limit for initial review