-- XML Index Database Schema
-- Run this on your SQL Server to create the database structure

USE master;
GO

-- Create database if it doesn't exist
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'XMLIndex')
BEGIN
    CREATE DATABASE XMLIndex;
END
GO

USE XMLIndex;
GO

-- Table to track XML files and their metadata
CREATE TABLE XMLFiles (
    ID int IDENTITY(1,1) PRIMARY KEY,
    FilePath nvarchar(500) NOT NULL,
    FileName nvarchar(255) NOT NULL,
    PartNumber nvarchar(100) NOT NULL,
    Revision nvarchar(10) NOT NULL,
    Release nvarchar(10) NOT NULL,
    FileModifiedDate datetime,
    ParsedDate datetime NOT NULL,
    UNIQUE(FilePath)
);

-- Table for part manufacturing data
CREATE TABLE PartData (
    ID int IDENTITY(1,1) PRIMARY KEY,
    XMLFileID int NOT NULL,
    PartNumber nvarchar(100) NOT NULL,
    Revision nvarchar(10) NOT NULL,
    Release nvarchar(10) NOT NULL,
    Description nvarchar(500),
    MakeBuy nvarchar(50),
    Material nvarchar(100),
    Thickness decimal(10,4),
    Weight decimal(10,4),
    MaxX decimal(10,4),
    MaxY decimal(10,4),
    MaxZ decimal(10,4),
    Rotation int,
    GangQty int,
    RawMaterialNumber nvarchar(100),
    FOREIGN KEY (XMLFileID) REFERENCES XMLFiles(ID)
);

-- Table for manufacturing flags
CREATE TABLE ManufacturingFlags (
    ID int IDENTITY(1,1) PRIMARY KEY,
    XMLFileID int NOT NULL,
    PartNumber nvarchar(100) NOT NULL,
    Laser bit DEFAULT 0,
    Punch bit DEFAULT 0,
    Saw bit DEFAULT 0,
    Shear bit DEFAULT 0,
    Powder bit DEFAULT 0,
    LoosePart bit DEFAULT 0,
    ShipLoose bit DEFAULT 0,
    AssemblyCut bit DEFAULT 0,
    HardwareLot bit DEFAULT 0,
    Template bit DEFAULT 0,
    TemplateCut bit DEFAULT 0,
    FOREIGN KEY (XMLFileID) REFERENCES XMLFiles(ID)
);

-- Table for BOM structure (assembly hierarchy)
CREATE TABLE BOMItems (
    ID int IDENTITY(1,1) PRIMARY KEY,
    ParentXMLFileID int NOT NULL,
    ParentPartNumber nvarchar(100) NOT NULL,
    ChildPartNumber nvarchar(100) NOT NULL,
    ChildRevision nvarchar(10),
    Quantity decimal(10,4),
    Level int,
    FOREIGN KEY (ParentXMLFileID) REFERENCES XMLFiles(ID)
);

-- Indexes for fast querying
CREATE INDEX IX_XMLFiles_PartNumber ON XMLFiles(PartNumber, Revision, Release);
CREATE INDEX IX_PartData_PartNumber ON PartData(PartNumber, Revision, Release);
CREATE INDEX IX_PartData_Material ON PartData(Material, Thickness);
CREATE INDEX IX_BOMItems_Parent ON BOMItems(ParentPartNumber);
CREATE INDEX IX_BOMItems_Child ON BOMItems(ChildPartNumber);

-- View for easy querying
CREATE VIEW vw_LatestParts AS
SELECT 
    xf.PartNumber,
    MAX(CAST(xf.Release AS int)) as LatestRelease,
    xf.Revision
FROM XMLFiles xf
GROUP BY xf.PartNumber, xf.Revision;

-- View for complete part details
CREATE VIEW vw_PartDetails AS
SELECT 
    xf.PartNumber,
    xf.Revision,
    xf.Release,
    xf.FileName,
    xf.FilePath,
    pd.Description,
    pd.MakeBuy,
    pd.Material,
    pd.Thickness,
    pd.Weight,
    pd.MaxX,
    pd.MaxY,
    pd.Rotation,
    pd.GangQty,
    pd.RawMaterialNumber,
    mf.Laser,
    mf.Punch,
    mf.Saw,
    mf.Shear,
    mf.Powder,
    mf.LoosePart,
    mf.ShipLoose,
    mf.Template,
    xf.ParsedDate,
    CASE WHEN lp.LatestRelease = CAST(xf.Release AS int) THEN 1 ELSE 0 END as IsLatestRelease
FROM XMLFiles xf
LEFT JOIN PartData pd ON xf.ID = pd.XMLFileID
LEFT JOIN ManufacturingFlags mf ON xf.ID = mf.XMLFileID
LEFT JOIN vw_LatestParts lp ON xf.PartNumber = lp.PartNumber AND xf.Revision = lp.Revision;

GO

-- Sample queries to test the schema
/*
-- Find latest release for a part
SELECT * FROM vw_PartDetails 
WHERE PartNumber = 'L156378' AND IsLatestRelease = 1;

-- Find parts by material and thickness
SELECT * FROM vw_PartDetails 
WHERE Material = 'Mild Steel (P&O)' AND Thickness = 0.0598;

-- Find similar sized parts
SELECT * FROM vw_PartDetails 
WHERE MaxX BETWEEN 20 AND 25 AND MaxY BETWEEN 15 AND 20;

-- Get BOM for an assembly
SELECT * FROM BOMItems 
WHERE ParentPartNumber = 'FP-550-1732-1' 
ORDER BY Level, ChildPartNumber;

-- Count total XMLs processed
SELECT COUNT(*) as TotalXMLs FROM XMLFiles;
*/
