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