-- Test queries for the XML Index database
-- Run these to verify the data was loaded correctly

-- 1. Basic counts
SELECT 
    COUNT(*) as TotalXMLFiles,
    COUNT(DISTINCT PartNumber) as UniqueParts,
    MIN(ProcessedDate) as FirstProcessed,
    MAX(ProcessedDate) as LastProcessed
FROM XMLFiles;

-- 2. Sample part data
SELECT 
    xf.PartNumber,
    xf.Revision,
    pd.Material,
    pd.Thickness,
    pd.Weight,
    xf.FilePath
FROM XMLFiles xf
LEFT JOIN PartData pd ON xf.ID = pd.XMLFileID
LIMIT 10;

-- 3. Most common materials
SELECT 
    Material,
    COUNT(*) as PartCount
FROM PartData 
WHERE Material IS NOT NULL AND Material != ''
GROUP BY Material
ORDER BY PartCount DESC
LIMIT 10;

-- 4. Recent files by part number pattern
SELECT 
    PartNumber,
    Revision,
    FilePath,
    ProcessedDate
FROM XMLFiles 
WHERE PartNumber LIKE 'HDW%'
ORDER BY ProcessedDate DESC
LIMIT 5;

-- 5. BOM relationships
SELECT 
    xf.PartNumber as ParentPart,
    bi.ChildPartNumber,
    bi.Quantity,
    COUNT(*) as RelationshipCount
FROM XMLFiles xf
JOIN BomItems bi ON xf.ID = bi.XMLFileID
WHERE bi.ChildPartNumber IS NOT NULL AND bi.ChildPartNumber != ''
GROUP BY xf.PartNumber, bi.ChildPartNumber, bi.Quantity
ORDER BY RelationshipCount DESC
LIMIT 10;