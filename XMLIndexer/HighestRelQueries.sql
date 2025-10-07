-- Highest RELXX Release Only - SULL-1006-0628 Make Items
-- This query extracts REL numbers and shows only components from the highest REL value

-- Method 1: Extract REL numbers and get highest for each base part
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
    'Highest REL Make Items for SULL-1006-0628' as QueryType,
    c.ComponentPartNumber, 
    c.ComponentType,
    c.TotalQuantity,
    pr.PartNumber as XMLPartNumber,
    pr.ReleaseNumber,
    pr.FilePath
FROM Components c
JOIN ParsedReleases pr ON c.XMLFileID = pr.ID
JOIN HighestReleases hr ON pr.BasePartNumber = hr.BasePartNumber 
    AND pr.ReleaseNumber = hr.HighestRelease
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber;

-- Method 2: Simpler regex-like approach for REL extraction
SELECT DISTINCT
    'Simple Highest REL Make Items' as QueryType,
    c.ComponentPartNumber, 
    c.ComponentType,
    c.TotalQuantity,
    xf.PartNumber as XMLPartNumber,
    -- Extract and show the REL number
    CASE 
        WHEN xf.PartNumber LIKE '%_REL%' THEN 
            CAST(SUBSTR(xf.PartNumber, INSTR(xf.PartNumber, '_REL') + 4) as INTEGER)
        ELSE 0
    END as ReleaseNumber
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
    AND xf.PartNumber = (
        -- Get the XML with highest REL number for this component
        SELECT xf2.PartNumber 
        FROM Components c2 
        JOIN XMLFiles xf2 ON c2.XMLFileID = xf2.ID
        WHERE c2.ComponentPartNumber = c.ComponentPartNumber 
            AND c2.ParentPartNumber LIKE '%SULL-1006-0628%'
            AND c2.ComponentType = 'Make'
            AND xf2.PartNumber LIKE '%SULL-1006-0628%'
        ORDER BY 
            -- Sort by REL number (highest first)
            CASE 
                WHEN xf2.PartNumber LIKE '%_REL%' THEN 
                    CAST(SUBSTR(xf2.PartNumber, INSTR(xf2.PartNumber, '_REL') + 4) as INTEGER)
                ELSE 0
            END DESC
        LIMIT 1
    )
ORDER BY c.ComponentPartNumber;

-- Method 3: Show all REL versions to verify which is highest
SELECT DISTINCT
    'All REL Versions for SULL-1006-0628' as QueryType,
    xf.PartNumber as XMLPartNumber,
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
GROUP BY xf.PartNumber
ORDER BY ReleaseNumber DESC;