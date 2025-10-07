-- Test H1319-0000 Workflow Query
-- Check if our test job's XML file exists in the database

-- 1. Check if SULL-1006-0627 XML file exists
SELECT 
    'XML Files for Test Part' as QueryType,
    xf.PartNumber,
    xf.Revision,
    xf.Release,
    xf.FilePath
FROM XMLFiles xf
WHERE xf.PartNumber LIKE '%SULL-1006-0627%'
ORDER BY xf.Release DESC;

-- 2. Check for components that might match our test job
SELECT 
    'Components for Test Part' as QueryType,
    c.ComponentPartNumber,
    c.ComponentType,
    c.TotalQuantity,
    c.Material,
    c.Thickness,
    xf.PartNumber as XMLFile
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
WHERE xf.PartNumber LIKE '%SULL-1006-0627%'
   OR c.ParentPartNumber LIKE '%SULL-1006-0627%'
   OR c.ComponentPartNumber LIKE '%SULL-1006-0627%'
ORDER BY c.ComponentPartNumber;

-- 3. Check MRP view for our test job
SELECT 
    'MRP View for Test Job' as QueryType,
    *
FROM vw_MrpWithXmlStatus 
WHERE JobNumber = 'H1319-0000' 
   OR PartNumber LIKE '%SULL-1006-0627%';

-- 4. Show all available XML files (sample)
SELECT 
    'Available XML Files Sample' as QueryType,
    PartNumber,
    Revision,
    Release,
    COUNT(*) as ComponentCount
FROM XMLFiles xf
JOIN Components c ON xf.ID = c.XMLFileID
GROUP BY xf.PartNumber, xf.Revision, xf.Release
ORDER BY xf.Release DESC
LIMIT 10;