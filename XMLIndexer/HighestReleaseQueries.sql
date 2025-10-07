-- Highest Release Only - SULL-1006-0628 Make Items
-- This query shows only components from the latest/highest release of each XML file

-- Method 1: Latest release by filtering highest revision in filename
WITH LatestXMLFiles AS (
    -- Get the highest release/revision for each base part number
    SELECT 
        -- Extract base part number (everything before _REV or _REL)
        CASE 
            WHEN PartNumber LIKE '%_REV%' THEN SUBSTR(PartNumber, 1, INSTR(PartNumber, '_REV') - 1)
            WHEN PartNumber LIKE '%_REL%' THEN SUBSTR(PartNumber, 1, INSTR(PartNumber, '_REL') - 1)
            ELSE PartNumber 
        END as BasePartNumber,
        MAX(PartNumber) as LatestPartNumber,  -- Assumes alphanumeric sorting gives latest
        MAX(ID) as LatestXMLFileID
    FROM XMLFiles 
    WHERE PartNumber LIKE '%SULL-1006-0628%'
    GROUP BY 
        CASE 
            WHEN PartNumber LIKE '%_REV%' THEN SUBSTR(PartNumber, 1, INSTR(PartNumber, '_REV') - 1)
            WHEN PartNumber LIKE '%_REL%' THEN SUBSTR(PartNumber, 1, INSTR(PartNumber, '_REL') - 1)
            ELSE PartNumber 
        END
)
SELECT DISTINCT
    'Latest Release Make Items for SULL-1006-0628' as QueryType,
    c.ComponentPartNumber, 
    c.ComponentType,
    c.TotalQuantity,
    xf.PartNumber as XMLPartNumber,
    xf.FilePath
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
JOIN LatestXMLFiles latest ON xf.ID = latest.LatestXMLFileID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
ORDER BY c.ComponentPartNumber;

-- Method 2: Simpler approach - just get max XMLFileID for each component
SELECT DISTINCT
    'Highest XMLFileID Make Items for SULL-1006-0628' as QueryType,
    c.ComponentPartNumber, 
    c.ComponentType,
    c.TotalQuantity,
    xf.PartNumber as XMLPartNumber
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
    AND c.XMLFileID = (
        -- Get the highest XMLFileID for this component
        SELECT MAX(c2.XMLFileID) 
        FROM Components c2 
        WHERE c2.ComponentPartNumber = c.ComponentPartNumber 
            AND c2.ParentPartNumber LIKE '%SULL-1006-0628%'
            AND c2.ComponentType = 'Make'
    )
ORDER BY c.ComponentPartNumber;

-- Method 3: Most recent by file timestamp
SELECT DISTINCT
    'Most Recent File Make Items for SULL-1006-0628' as QueryType,
    c.ComponentPartNumber, 
    c.ComponentType,
    c.TotalQuantity,
    xf.PartNumber as XMLPartNumber,
    xf.LastModified
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make' 
    AND xf.LastModified = (
        -- Get the most recent file for this component
        SELECT MAX(xf2.LastModified) 
        FROM Components c2 
        JOIN XMLFiles xf2 ON c2.XMLFileID = xf2.ID
        WHERE c2.ComponentPartNumber = c.ComponentPartNumber 
            AND c2.ParentPartNumber LIKE '%SULL-1006-0628%'
            AND c2.ComponentType = 'Make'
    )
ORDER BY c.ComponentPartNumber;