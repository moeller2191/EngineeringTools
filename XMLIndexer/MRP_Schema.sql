-- Add MRP Priority List table to XMLIndex database
-- This table will store the current priority list data from the Excel file

-- Drop existing table if it exists
DROP TABLE IF EXISTS MrpPriorityList;

-- Create MRP Priority List table
CREATE TABLE MrpPriorityList (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    JobNumber TEXT NOT NULL,
    PartNumber TEXT,
    Revision TEXT,
    Quantity INTEGER DEFAULT 1,
    Description TEXT,
    Priority INTEGER,
    DueDate TEXT,
    Status TEXT,
    Customer TEXT,
    Program TEXT,
    Notes TEXT,
    LastUpdated DATETIME DEFAULT CURRENT_TIMESTAMP,
    
    -- Create index for fast lookups
    UNIQUE(JobNumber, PartNumber, Revision)
);

-- Create indexes for performance
CREATE INDEX idx_mrp_job_number ON MrpPriorityList(JobNumber);
CREATE INDEX idx_mrp_part_number ON MrpPriorityList(PartNumber);
CREATE INDEX idx_mrp_priority ON MrpPriorityList(Priority);
CREATE INDEX idx_mrp_status ON MrpPriorityList(Status);

INSERT INTO MrpPriorityList (JobNumber, PartNumber, Revision, Quantity, Description, Priority, Status) VALUES
('G0770-0000', 'SULL-I-02250132-241', 'REV01', 1, 'Sullivan Assembly - Sample Part', 1, 'Active'),
('G0783-0000', 'SPI-01901000-0943GRAY', 'REV05', 2, 'SPI Gray Component', 2, 'Active'),
('G0809-0000', 'SPI-01901000-1161GT', 'REV01', 1, 'SPI GT Component', 3, 'Active'),
('G4470-0000', 'SULL-02250157-350', 'REV03', 3, 'Sullivan 350 Series', 1, 'Active'),
('G0840-0000', 'SPI-03903297-0030WM', 'REV01', 1, 'SPI WM Series Part', 2, 'Active'),
('G0841-0000', 'SPI-03903297-0031WM', 'REV01', 1, 'SPI WM Series Part', 2, 'Active'),
('G0842-0000', 'SPI-03903297-0032WM', 'REV01', 2, 'SPI WM Series Part', 3, 'Active'),
('G0843-0000', 'SPI-03903297-0033WM', 'REV01', 1, 'SPI WM Series Part', 3, 'Active'),
('G0818-0000', 'SPI-01901000-1013GRAY', 'REV00', 1, 'SPI Gray Component', 1, 'Active'),
('G0819-0000', 'SPI-01901000-1014GRAY', 'REV01', 1, 'SPI Gray Component', 1, 'Active'),
('TEST-0001', 'SULL-1006-0628', 'A', 1, 'Test Assembly for Demo', 1, 'Active'),
('IK3NC-0000', 'SULL-I-02250180-560', 'REV02', 2, 'Sullivan I-Job Assembly - Structural Component', 1, 'Active'),
('IK3NC-0001', 'SPI-I-01901000-1050GRAY', 'REV01', 1, 'SPI I-Job Gray Component - Custom Fabrication', 2, 'Active'),
('IK3NC-0002', 'SULL-I-02250252-633', 'REV03', 3, 'Sullivan I-Job Door Assembly - Custom Build', 1, 'Active'),
('IL7MP-0000', 'SPI-I-03903297-0040WM', 'REV01', 1, 'SPI I-Job WM Series - Manufacturing Part', 3, 'Active'),
('IL7MP-0001', 'SULL-I-02250157-420', 'REV02', 2, 'Sullivan I-Job 420 Series - Internal Component', 2, 'Active'),
('IM9QR-0000', 'SPI-I-20144-256', 'REV01', 5, 'SPI I-Job Multi-Component Assembly', 1, 'Active');

-- Add the requested test job for workflow validation
INSERT INTO MrpPriorityList (JobNumber, PartNumber, Revision, Quantity, Description, Priority, Status) VALUES
('H1319-0000', 'SULL-1006-0627', '', 10, 'Test Job for Workflow', 1, 'Active');

-- Create a view for easy MRP data access with XML status
CREATE VIEW vw_MrpWithXmlStatus AS
SELECT 
    m.ID,
    m.JobNumber,
    m.PartNumber,
    m.Revision,
    m.Quantity,
    m.Description,
    m.Priority,
    m.DueDate,
    m.Status,
    m.Customer,
    m.Program,
    m.Notes,
    m.LastUpdated,
    
    -- Check if XML exists for this part
    CASE 
        WHEN xf.ID IS NOT NULL THEN 'Available'
        ELSE 'Not Found'
    END as XmlStatus,
    
    -- Get highest release if XML exists
    COALESCE(MAX(CAST(xf.Release AS INTEGER)), 0) as HighestRelease,
    
    -- Count of components if XML exists
    COUNT(c.ID) as ComponentCount
    
FROM MrpPriorityList m
LEFT JOIN XMLFiles xf ON (
    xf.PartNumber LIKE '%' || m.PartNumber || '%' OR
    m.PartNumber LIKE '%' || REPLACE(xf.PartNumber, '_REV', '') || '%'
)
LEFT JOIN Components c ON xf.ID = c.XMLFileID AND c.ComponentType = 'Make'
GROUP BY m.ID, m.JobNumber, m.PartNumber, m.Revision
ORDER BY m.Priority, m.JobNumber;