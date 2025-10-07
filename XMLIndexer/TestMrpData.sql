-- Test MRP Integration
-- This script sets up and tests the MRP functionality

-- First, ensure the MRP table exists and has sample data
INSERT OR IGNORE INTO MrpPriorityList (JobNumber, PartNumber, Revision, Quantity, Description, Priority, Status) VALUES
('G0770-0000', 'SULL-I-02250132-241', 'REV01', 1, 'Sullivan Assembly - Sample Part', 1, 'Active'),
('G0783-0000', 'SPI-01901000-0943GRAY', 'REV05', 2, 'SPI Gray Component', 2, 'Active'),
('G0809-0000', 'SPI-01901000-1161GT', 'REV01', 1, 'SPI GT Component', 3, 'Active'),
('G4470-0000', 'SULL-02250157-350', 'REV03', 3, 'Sullivan 350 Series', 1, 'Active'),
('TEST-0001', 'SULL-1006-0628', 'A', 1, 'Test Assembly for Demo', 1, 'Active');

-- Test query: Show MRP data with XML status
SELECT 'MRP Data with XML Status:' as Test;
SELECT 
    JobNumber,
    PartNumber,
    Revision,
    Quantity,
    Description,
    XmlStatus,
    HighestRelease,
    ComponentCount
FROM vw_MrpWithXmlStatus
ORDER BY Priority, JobNumber;

-- Test query: Show specific job data
SELECT 'Job TEST-0001 Details:' as Test;
SELECT * FROM vw_MrpWithXmlStatus WHERE JobNumber = 'TEST-0001';

-- Test query: Show all active jobs
SELECT 'All Active Jobs:' as Test;
SELECT JobNumber, PartNumber, Description, XmlStatus FROM vw_MrpWithXmlStatus WHERE Status = 'Active';

-- Test query: Show XML files that match our MRP parts
SELECT 'XML Files matching MRP parts:' as Test;
SELECT DISTINCT
    m.JobNumber,
    m.PartNumber as MRP_PartNumber,
    xf.PartNumber as XML_PartNumber,
    xf.Release,
    COUNT(c.ID) as ComponentCount
FROM MrpPriorityList m
LEFT JOIN XMLFiles xf ON (
    xf.PartNumber LIKE '%' || m.PartNumber || '%' OR
    m.PartNumber LIKE '%' || REPLACE(xf.PartNumber, '_REV', '') || '%'
)
LEFT JOIN Components c ON xf.ID = c.XMLFileID AND c.ComponentType = 'Make'
WHERE m.Status = 'Active'
GROUP BY m.JobNumber, m.PartNumber, xf.PartNumber, xf.Release
ORDER BY m.JobNumber, xf.Release DESC;