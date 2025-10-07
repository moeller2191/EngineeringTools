-- CHECK: Do individual component XML files exist in our database?
-- Let's see if we have XMLFiles for the individual components

SELECT 
    'Individual Component XMLFiles Check' as CheckType,
    c.ComponentPartNumber,
    xf_comp.PartNumber as ComponentXMLFile,
    xf_comp.Release as ComponentRelease,
    pd.MaxX,
    pd.MaxY,
    pd.RawMaterialNumber
FROM Components c
JOIN XMLFiles xf ON c.XMLFileID = xf.ID
LEFT JOIN XMLFiles xf_comp ON (
    xf_comp.PartNumber = c.ComponentPartNumber OR
    xf_comp.PartNumber = REPLACE(c.ComponentPartNumber, '.SLDPRT', '') OR
    xf_comp.PartNumber LIKE '%' || REPLACE(c.ComponentPartNumber, '.SLDPRT', '') || '%'
)
LEFT JOIN PartData pd ON pd.XMLFileID = xf_comp.ID
WHERE c.ParentPartNumber LIKE '%SULL-1006-0628%' 
    AND c.ComponentType = 'Make'
    AND xf.Release = '53'
ORDER BY c.ComponentPartNumber;

-- CHECK: What XMLFiles do we actually have that start with L15?
SELECT 
    'Available L15 XMLFiles' as CheckType,
    PartNumber,
    Release,
    FilePath
FROM XMLFiles 
WHERE PartNumber LIKE 'L15%'
ORDER BY PartNumber
LIMIT 20;

-- CHECK: Do we have PartData for any L15 parts?
SELECT 
    'Available L15 PartData' as CheckType,
    pd.PartNumber,
    pd.Description,
    pd.MaxX,
    pd.MaxY,
    pd.RawMaterialNumber,
    xf.PartNumber as XMLFile
FROM PartData pd
JOIN XMLFiles xf ON pd.XMLFileID = xf.ID
WHERE pd.PartNumber LIKE 'L15%'
ORDER BY pd.PartNumber
LIMIT 20;