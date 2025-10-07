-- Add I-job samples to MRP database for testing
-- This script adds sample I-job entries to test I-job functionality

INSERT INTO MrpPriorityList (JobNumber, PartNumber, Revision, Quantity, Description, Priority, Status) VALUES
('IK3NC-0000', 'SULL-I-02250180-560', 'REV02', 2, 'Sullivan I-Job Assembly - Structural Component', 1, 'Active'),
('IK3NC-0001', 'SPI-I-01901000-1050GRAY', 'REV01', 1, 'SPI I-Job Gray Component - Custom Fabrication', 2, 'Active'),
('IK3NC-0002', 'SULL-I-02250252-633', 'REV03', 3, 'Sullivan I-Job Door Assembly - Custom Build', 1, 'Active'),
('IL7MP-0000', 'SPI-I-03903297-0040WM', 'REV01', 1, 'SPI I-Job WM Series - Manufacturing Part', 3, 'Active'),
('IL7MP-0001', 'SULL-I-02250157-420', 'REV02', 2, 'Sullivan I-Job 420 Series - Internal Component', 2, 'Active'),
('IM9QR-0000', 'SPI-I-20144-256', 'REV01', 5, 'SPI I-Job Multi-Component Assembly', 1, 'Active');