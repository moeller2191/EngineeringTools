using System;
using System.IO;
using System.Xml;
using Microsoft.Data.Sqlite;
using System.Collections.Generic;
using System.Linq;

namespace XMLIndexer
{
    class Program
    {
        // Configuration - update these for your environment
        private static readonly string[] XmlPaths = {
            @"\\kmi-solidworks22\solidworks22common\CUT LIST XML",
            @"\\kmi-solidworks22\solidworks22common\CUT LIST XML\Legacy", 
            @"\\kmi-solidworks22\solidworks22common\CUT LIST XML\New"
        };
        
        // SQLite database file - will be created automatically
        private static readonly string DatabasePath = "XMLIndex.db";
        private static readonly string ConnectionString = $"Data Source={DatabasePath}";

        static void Main(string[] args)
        {
            Console.WriteLine("=== XML INDEX BUILDER ===");
            
            // Parse command line arguments
            bool forceFullScan = args.Contains("--full") || args.Contains("-f");
            bool incrementalOnly = args.Contains("--incremental") || args.Contains("-i");
            bool showHelp = args.Contains("--help") || args.Contains("-h");
            bool exploreDatabase = args.Contains("--explore") || args.Contains("-e");
            bool searchMode = args.Contains("--search") || args.Contains("-s");
            bool testMrp = args.Contains("--test-mrp") || args.Contains("-t");
            bool applySchema = args.Contains("--apply-schema") || args.Contains("-a");
            
            if (applySchema)
            {
                ApplyMrpSchema();
                return;
            }
            
            if (testMrp)
            {
                MrpTester.TestMrpFunctionality(DatabasePath);
                return;
            }
            
            if (showHelp)
            {
                ShowHelp();
                return;
            }
            
            if (exploreDatabase)
            {
                ExploreDatabase();
                return;
            }
            
            if (searchMode)
            {
                SearchDatabase(args);
                return;
            }
            
            Console.WriteLine("Processing your XML files...");
            Console.WriteLine($"Database will be created at: {Path.GetFullPath(DatabasePath)}");
            
            if (forceFullScan)
                Console.WriteLine("Mode: FULL SCAN (will reprocess all files)");
            else if (incrementalOnly)
                Console.WriteLine("Mode: INCREMENTAL ONLY (new/modified files only)");
            else
                Console.WriteLine("Mode: SMART UPDATE (new files + test mode)");
                
            Console.WriteLine();
            
            try
            {
                // Create database and tables
                CreateDatabase();
                
                // Scan for ALL XML files
                var xmlFiles = ScanAllXmlFiles();
                Console.WriteLine($"Found {xmlFiles.Count} XML files total");
                
                if (xmlFiles.Count == 0)
                {
                    Console.WriteLine("No XML files found. Check your paths:");
                    foreach (var path in XmlPaths)
                    {
                        Console.WriteLine($"  {path}");
                    }
                    return;
                }
                
                // Get files that need processing (new or modified)
                var filesToProcess = GetFilesToProcess(xmlFiles, forceFullScan);
                Console.WriteLine($"Files needing processing: {filesToProcess.Count}");
                
                if (filesToProcess.Count == 0)
                {
                    Console.WriteLine("✅ All files are up to date!");
                    ShowProcessingSummary();
                    ShowSampleQueries();
                    return;
                }
                
                // Process files based on mode
                Console.WriteLine();
                
                if (incrementalOnly)
                {
                    // Incremental mode - just process the files that need updating
                    Console.WriteLine($"Processing {filesToProcess.Count} new/modified files...");
                    ProcessXmlFiles(filesToProcess);
                }
                else if (forceFullScan)
                {
                    // Full scan mode - process everything
                    Console.WriteLine($"Processing all {filesToProcess.Count} files...");
                    ProcessXmlFiles(filesToProcess);
                }
                else
                {
                    // Smart mode - test first, then ask about remaining
                    if (filesToProcess.Count <= 10)
                    {
                        Console.WriteLine($"Processing {filesToProcess.Count} new/modified files...");
                        ProcessXmlFiles(filesToProcess);
                    }
                    else
                    {
                        Console.WriteLine("Processing first 10 new/modified files as a test...");
                        var testFiles = filesToProcess.Take(10).ToList();
                        ProcessXmlFiles(testFiles);
                        
                        // Ask if user wants to continue with remaining files
                        var remainingFiles = filesToProcess.Skip(10).ToList();
                        if (remainingFiles.Count > 0)
                        {
                            Console.WriteLine();
                            Console.Write($"Test complete! Process remaining {remainingFiles.Count} files? (y/n): ");
                            var response = Console.ReadLine();
                            
                            if (response?.ToLower().StartsWith("y") == true)
                            {
                                ProcessXmlFiles(remainingFiles);
                            }
                        }
                    }
                }
                
                // Show summary
                ShowProcessingSummary();
                
                // Show sample queries
                ShowSampleQueries();
                
                Console.WriteLine();
                Console.WriteLine("=== PROCESSING COMPLETE ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FATAL ERROR: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
        
        static void CreateDatabase()
        {
            Console.WriteLine("Setting up SQLite database...");
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var createTables = @"
                -- Table to track XML files and their metadata
                CREATE TABLE IF NOT EXISTS XMLFiles (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    FilePath TEXT NOT NULL,
                    FileName TEXT NOT NULL,
                    PartNumber TEXT NOT NULL,
                    Revision TEXT NOT NULL,
                    Release TEXT NOT NULL,
                    FileModifiedDate TEXT,
                    ParsedDate TEXT NOT NULL,
                    UNIQUE(FilePath)
                );

                -- Table for part manufacturing data
                CREATE TABLE IF NOT EXISTS PartData (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    XMLFileID INTEGER NOT NULL,
                    PartNumber TEXT NOT NULL,
                    Revision TEXT NOT NULL,
                    Release TEXT NOT NULL,
                    Description TEXT,
                    MakeBuy TEXT,
                    Material TEXT,
                    Thickness REAL,
                    Weight REAL,
                    MaxX REAL,
                    MaxY REAL,
                    MaxZ REAL,
                    Rotation INTEGER,
                    GangQty INTEGER,
                    RawMaterialNumber TEXT,
                    FOREIGN KEY (XMLFileID) REFERENCES XMLFiles(ID)
                );

                -- Table for manufacturing flags
                CREATE TABLE IF NOT EXISTS ManufacturingFlags (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    XMLFileID INTEGER NOT NULL,
                    PartNumber TEXT NOT NULL,
                    Laser INTEGER DEFAULT 0,
                    Punch INTEGER DEFAULT 0,
                    Saw INTEGER DEFAULT 0,
                    Shear INTEGER DEFAULT 0,
                    Powder INTEGER DEFAULT 0,
                    LoosePart INTEGER DEFAULT 0,
                    ShipLoose INTEGER DEFAULT 0,
                    AssemblyCut INTEGER DEFAULT 0,
                    HardwareLot INTEGER DEFAULT 0,
                    Template INTEGER DEFAULT 0,
                    TemplateCut INTEGER DEFAULT 0,
                    FOREIGN KEY (XMLFileID) REFERENCES XMLFiles(ID)
                );

                -- Table for component/make items
                CREATE TABLE IF NOT EXISTS Components (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    XMLFileID INTEGER NOT NULL,
                    ParentPartNumber TEXT NOT NULL,
                    ComponentPartNumber TEXT,
                    ComponentDescription TEXT,
                    ComponentType TEXT, -- 'make', 'buy', 'weldment', etc.
                    Quantity INTEGER DEFAULT 1,
                    TotalQuantity INTEGER DEFAULT 1, -- Calculated quantity considering parent assemblies
                    Material TEXT,
                    Thickness REAL,
                    ComponentIndex INTEGER, -- order in the BOM
                    AssemblyLevel INTEGER DEFAULT 1, -- depth in assembly hierarchy
                    -- Enhanced PartData properties for components
                    Weight REAL,
                    MaxX REAL,
                    MaxY REAL,
                    MaxZ REAL,
                    Rotation INTEGER,
                    GangQty INTEGER,
                    RawMaterialNumber TEXT,
                    FOREIGN KEY (XMLFileID) REFERENCES XMLFiles(ID)
                );

                -- Indexes for fast querying
                CREATE INDEX IF NOT EXISTS IX_XMLFiles_PartNumber ON XMLFiles(PartNumber, Revision, Release);
                CREATE INDEX IF NOT EXISTS IX_PartData_PartNumber ON PartData(PartNumber, Revision, Release);
                CREATE INDEX IF NOT EXISTS IX_PartData_Material ON PartData(Material, Thickness);
                CREATE INDEX IF NOT EXISTS IX_Components_Parent ON Components(ParentPartNumber, ComponentType);
                CREATE INDEX IF NOT EXISTS IX_Components_Component ON Components(ComponentPartNumber, ComponentType);
            ";
            
            var command = new SqliteCommand(createTables, connection);
            command.ExecuteNonQuery();
            
            Console.WriteLine("✓ Database created successfully!");
            Console.WriteLine();
        }
        
        static List<string> ScanAllXmlFiles()
        {
            Console.WriteLine("Scanning for XML files...");
            var xmlFiles = new List<string>();
            
            foreach (var path in XmlPaths)
            {
                Console.WriteLine($"Scanning: {path}");
                try
                {
                    if (Directory.Exists(path))
                    {
                        var files = Directory.GetFiles(path, "*.xml", SearchOption.TopDirectoryOnly);
                        xmlFiles.AddRange(files);
                        Console.WriteLine($"  Found {files.Length} files");
                    }
                    else
                    {
                        Console.WriteLine($"  ⚠ Path not accessible: {path}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  ✗ Error scanning {path}: {ex.Message}");
                }
            }
            
            Console.WriteLine($"Total XML files found: {xmlFiles.Count}");
            Console.WriteLine();
            return xmlFiles;
        }
        
        static void ProcessXmlFiles(List<string> xmlFiles)
        {
            Console.WriteLine($"Processing {xmlFiles.Count} XML files...");
            var processedCount = 0;
            var errorCount = 0;
            var startTime = DateTime.Now;
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            foreach (var filePath in xmlFiles)
            {
                try
                {
                    // Check if already processed
                    if (IsFileAlreadyProcessed(connection, filePath))
                    {
                        continue;
                    }
                    
                    var xmlData = ParseXmlFile(filePath);
                    if (xmlData != null)
                    {
                        InsertXmlData(connection, xmlData);
                        processedCount++;
                        
                        // Progress indicator
                        if (processedCount % 10 == 0)
                        {
                            var elapsed = DateTime.Now - startTime;
                            var rate = processedCount / Math.Max(elapsed.TotalMinutes, 0.1);
                            Console.WriteLine($"Processed {processedCount}/{xmlFiles.Count} files ({rate:F1} files/min)");
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }
            
            Console.WriteLine();
            Console.WriteLine($"Processing complete! Processed: {processedCount}, Errors: {errorCount}");
        }
        
        static bool IsFileAlreadyProcessed(SqliteConnection connection, string filePath)
        {
            var cmd = new SqliteCommand("SELECT COUNT(*) FROM XMLFiles WHERE FilePath = @path", connection);
            cmd.Parameters.AddWithValue("@path", filePath);
            var count = Convert.ToInt32(cmd.ExecuteScalar());
            return count > 0;
        }
        
        static XmlPartData? ParseXmlFile(string filePath)
        {
            var fileName = Path.GetFileName(filePath);
            var xmlData = new XmlPartData
            {
                FilePath = filePath,
                FileName = fileName,
                ParsedDate = DateTime.Now,
                FileModifiedDate = File.GetLastWriteTime(filePath)
            };
            
            // Extract part info from filename (PartNumber_REVXXX_RELXX.xml)
            if (!ParseFileNameComponents(fileName, xmlData))
            {
                return null; // Skip files that don't match naming convention
            }
            
            try
            {
                // Load and parse XML content
                var xmlDoc = new XmlDocument();
                xmlDoc.Load(filePath);
                
                // Find main document node
                var docNode = xmlDoc.SelectSingleNode("//xml/transactions/transaction/document");
                if (docNode != null)
                {
                    ParseXmlAttributes(docNode, xmlData);
                    ParseBomStructure(docNode, xmlData);
                }
                
                return xmlData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"XML parsing error for {fileName}: {ex.Message}");
                return null;
            }
        }
        
        static bool ParseFileNameComponents(string fileName, XmlPartData xmlData)
        {
            try
            {
                // Parse: PartNumber_REVXXX_RELXX.xml
                if (fileName.Contains("_REV") && fileName.Contains("_REL"))
                {
                    var revIndex = fileName.IndexOf("_REV");
                    var relIndex = fileName.IndexOf("_REL");
                    var dotIndex = fileName.LastIndexOf('.');
                    
                    xmlData.PartNumber = fileName.Substring(0, revIndex);
                    xmlData.Revision = fileName.Substring(revIndex + 4, 3);
                    xmlData.Release = fileName.Substring(relIndex + 4, dotIndex - relIndex - 4);
                    
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        static void ParseXmlAttributes(XmlNode docNode, XmlPartData xmlData)
        {
            xmlData.MakeBuy = GetXmlAttribute(docNode, "PDMsmparttoggle");
            xmlData.Material = GetXmlAttribute(docNode, "PDMmatltype");
            xmlData.Thickness = ParseDecimal(GetXmlAttribute(docNode, "PDMthickness"));
            xmlData.Weight = ParseDecimal(GetXmlAttribute(docNode, "PDMpartweight"));
            xmlData.MaxX = ParseDecimal(GetXmlAttribute(docNode, "PDMmaxX"));
            xmlData.MaxY = ParseDecimal(GetXmlAttribute(docNode, "PDMmaxY"));
            xmlData.MaxZ = ParseDecimal(GetXmlAttribute(docNode, "PDMmaxZ"));
            xmlData.Description = GetXmlAttribute(docNode, "PDMchilddesc");
            xmlData.Rotation = ParseInt(GetXmlAttribute(docNode, "PDMrotation"));
            xmlData.GangQty = ParseInt(GetXmlAttribute(docNode, "PDMgangqty"));
            xmlData.RawMaterialNumber = GetXmlAttribute(docNode, "PDMrawmatlnum");
            
            // Manufacturing flags
            xmlData.Laser = GetXmlAttribute(docNode, "PDMlaser") == "1";
            xmlData.Punch = GetXmlAttribute(docNode, "PDMpunch") == "1";
            xmlData.Saw = GetXmlAttribute(docNode, "PDMsawcut") == "1";
            xmlData.Shear = GetXmlAttribute(docNode, "PDMshear") == "1";
            xmlData.Powder = GetXmlAttribute(docNode, "PDMpowdered") == "1";
            xmlData.LoosePart = GetXmlAttribute(docNode, "PDMlooseprt") == "1";
            xmlData.ShipLoose = GetXmlAttribute(docNode, "PDMshiploose") == "1";
            xmlData.AssemblyCut = GetXmlAttribute(docNode, "PDMassemblycut") == "1";
            xmlData.HardwareLot = GetXmlAttribute(docNode, "PDMhardwarelot") == "1";
            xmlData.Template = GetXmlAttribute(docNode, "PDMtemplate") == "1";
            xmlData.TemplateCut = GetXmlAttribute(docNode, "PDMtemplatecut") == "1";
        }
        
        static void ParseBomStructure(XmlNode docNode, XmlPartData xmlData)
        {
            int componentIndex = 0;
            
            // First, check if this assembly itself is marked as "Buy" or "Stock"
            // If so, we don't want to extract its components since we'll purchase the complete assembly
            var assemblyType = GetXmlAttribute(docNode, "PDMsmparttoggle");
            bool isAssemblyPurchased = !string.IsNullOrEmpty(assemblyType) && 
                                     (assemblyType.ToLower() == "buy" || assemblyType.ToLower() == "stock");
            
            if (isAssemblyPurchased)
            {
                Console.WriteLine($"DEBUG: Assembly {xmlData.PartNumber} is marked as '{assemblyType}' - skipping component extraction");
                return; // Don't extract components from purchased assemblies
            }
            
            // Strategy 1: Look for ALL nodes that have PDMsmparttoggle attribute anywhere in the document
            var allNodesWithType = docNode.SelectNodes(".//attribute[@name='PDMsmparttoggle']");
            if (allNodesWithType != null)
            {
                Console.WriteLine($"DEBUG: Found {allNodesWithType.Count} PDMsmparttoggle nodes in {xmlData.PartNumber}");
                foreach (XmlNode typeNode in allNodesWithType)
                {
                    var componentType = typeNode.Attributes?["value"]?.Value ?? "";
                    var parentElement = typeNode.ParentNode;
                    
                    if (parentElement != null)
                    {
                        // Skip if this component is also marked as "buy" or "stock" (nested purchased assemblies)
                        if (componentType.ToLower() == "buy" || componentType.ToLower() == "stock")
                        {
                            Console.WriteLine($"DEBUG: Skipping nested purchased assembly: Type='{componentType}'");
                            continue;
                        }
                        
                        // Try multiple ways to find the part number
                        var partNumber = parentElement.Attributes?["Name"]?.Value ??
                                       GetXmlAttribute(parentElement, "Name") ??
                                       GetXmlAttribute(parentElement, "PDMchildpartno") ??
                                       GetXmlAttribute(parentElement, "PartNumber") ??
                                       GetXmlAttribute(parentElement, "PDMpartno") ?? "";
                        
                        var description = GetXmlAttribute(parentElement, "PDMchilddesc") ??
                                        GetXmlAttribute(parentElement, "Description") ?? "";
                        
                        var material = GetXmlAttribute(parentElement, "PDMmatltype") ??
                                     GetXmlAttribute(parentElement, "Material") ?? "";
                        
                        var thickness = ParseDecimal(GetXmlAttribute(parentElement, "PDMthickness"));
                        
                        // Try to get quantity from various possible locations
                        var quantity = ParseDecimal(GetXmlAttribute(parentElement, "Quantity")) ??
                                     ParseDecimal(GetXmlAttribute(parentElement, "PDMquantity")) ??
                                     ParseDecimal(parentElement.Attributes?["quantity"]?.Value ?? "") ?? 1;
                        
                        // Extract all PartData properties for this component
                        var weight = ParseDecimal(GetXmlAttribute(parentElement, "PDMpartweight"));
                        var maxX = ParseDecimal(GetXmlAttribute(parentElement, "PDMmaxX"));
                        var maxY = ParseDecimal(GetXmlAttribute(parentElement, "PDMmaxY"));
                        var maxZ = ParseDecimal(GetXmlAttribute(parentElement, "PDMmaxZ"));
                        var rotation = ParseInt(GetXmlAttribute(parentElement, "PDMrotation"));
                        var gangQty = ParseInt(GetXmlAttribute(parentElement, "PDMgangqty"));
                        var rawMaterialNumber = GetXmlAttribute(parentElement, "PDMrawmatlnum") ?? "";
                        
                        Console.WriteLine($"DEBUG: Found component: Part='{partNumber}' Type='{componentType}' Desc='{description}'");
                        
                        var bomItem = new BomItem
                        {
                            ChildPartNumber = partNumber,
                            ComponentType = componentType,
                            Description = description,
                            Material = material,
                            Thickness = thickness,
                            Quantity = quantity,
                            TotalQuantity = quantity, // Initially set to base quantity
                            Level = 1,
                            ComponentIndex = componentIndex++,
                            AssemblyLevel = 1,
                            // Enhanced PartData properties
                            Weight = weight,
                            MaxX = maxX,
                            MaxY = maxY,
                            MaxZ = maxZ,
                            Rotation = rotation,
                            GangQty = gangQty,
                            RawMaterialNumber = rawMaterialNumber
                        };
                        
                        // Add if we have either a part number or component type
                        if (!string.IsNullOrEmpty(bomItem.ChildPartNumber) || !string.IsNullOrEmpty(bomItem.ComponentType))
                        {
                            xmlData.BomItems.Add(bomItem);
                        }
                    }
                }
            }
            
            // Strategy 2: Look for document nodes with attributes
            var documentNodes = docNode.SelectNodes(".//document");
            if (documentNodes != null)
            {
                Console.WriteLine($"DEBUG: Found {documentNodes.Count} document nodes in {xmlData.PartNumber}");
                foreach (XmlNode docElement in documentNodes)
                {
                    var componentType = GetXmlAttribute(docElement, "PDMsmparttoggle");
                    if (!string.IsNullOrEmpty(componentType))
                    {
                        var partNumber = docElement.Attributes?["Name"]?.Value ??
                                       GetXmlAttribute(docElement, "Name") ??
                                       GetXmlAttribute(docElement, "PDMchildpartno") ?? "";
                        
                        var description = GetXmlAttribute(docElement, "PDMchilddesc");
                        var material = GetXmlAttribute(docElement, "PDMmatltype");
                        var thickness = ParseDecimal(GetXmlAttribute(docElement, "PDMthickness"));
                        var quantity = ParseDecimal(GetXmlAttribute(docElement, "Quantity")) ?? 1;
                        
                        // Extract all PartData properties for this component
                        var weight = ParseDecimal(GetXmlAttribute(docElement, "PDMpartweight"));
                        var maxX = ParseDecimal(GetXmlAttribute(docElement, "PDMmaxX"));
                        var maxY = ParseDecimal(GetXmlAttribute(docElement, "PDMmaxY"));
                        var maxZ = ParseDecimal(GetXmlAttribute(docElement, "PDMmaxZ"));
                        var rotation = ParseInt(GetXmlAttribute(docElement, "PDMrotation"));
                        var gangQty = ParseInt(GetXmlAttribute(docElement, "PDMgangqty"));
                        var rawMaterialNumber = GetXmlAttribute(docElement, "PDMrawmatlnum") ?? "";
                        
                        Console.WriteLine($"DEBUG: Document component: Part='{partNumber}' Type='{componentType}'");
                        
                        var bomItem = new BomItem
                        {
                            ChildPartNumber = partNumber,
                            ComponentType = componentType,
                            Description = description,
                            Material = material,
                            Thickness = thickness,
                            Quantity = quantity,
                            TotalQuantity = quantity,
                            Level = 1,
                            ComponentIndex = componentIndex++,
                            AssemblyLevel = 1,
                            // Enhanced PartData properties
                            Weight = weight,
                            MaxX = maxX,
                            MaxY = maxY,
                            MaxZ = maxZ,
                            Rotation = rotation,
                            GangQty = gangQty,
                            RawMaterialNumber = rawMaterialNumber
                        };
                        
                        if (!string.IsNullOrEmpty(bomItem.ChildPartNumber) || !string.IsNullOrEmpty(bomItem.ComponentType))
                        {
                            xmlData.BomItems.Add(bomItem);
                        }
                    }
                }
            }
            
            // Strategy 3: Look for any element that contains "make", "buy", "stock" in attributes or text
            var makeNodes = docNode.SelectNodes(".//attribute[@value='make' or @value='Make' or @value='MAKE']");
            var buyNodes = docNode.SelectNodes(".//attribute[@value='buy' or @value='Buy' or @value='BUY']");
            var stockNodes = docNode.SelectNodes(".//attribute[@value='stock' or @value='Stock' or @value='STOCK']");
            
            var allMakeBuyNodes = new List<XmlNode>();
            if (makeNodes != null) allMakeBuyNodes.AddRange(makeNodes.Cast<XmlNode>());
            if (buyNodes != null) allMakeBuyNodes.AddRange(buyNodes.Cast<XmlNode>());
            if (stockNodes != null) allMakeBuyNodes.AddRange(stockNodes.Cast<XmlNode>());
            
            if (allMakeBuyNodes.Count > 0)
            {
                Console.WriteLine($"DEBUG: Found {allMakeBuyNodes.Count} make/buy/stock attributes in {xmlData.PartNumber}");
                foreach (XmlNode node in allMakeBuyNodes)
                {
                    var componentType = node.Attributes?["value"]?.Value ?? "";
                    var parentElement = node.ParentNode;
                    
                    if (parentElement != null)
                    {
                        var partNumber = parentElement.Attributes?["Name"]?.Value ??
                                       GetXmlAttribute(parentElement, "Name") ?? "";
                        
                        Console.WriteLine($"DEBUG: Make/Buy/Stock component: Part='{partNumber}' Type='{componentType}'");
                        
                        var bomItem = new BomItem
                        {
                            ChildPartNumber = partNumber,
                            ComponentType = componentType,
                            Description = GetXmlAttribute(parentElement, "PDMchilddesc"),
                            Material = GetXmlAttribute(parentElement, "PDMmatltype"),
                            Thickness = ParseDecimal(GetXmlAttribute(parentElement, "PDMthickness")),
                            Quantity = 1,
                            TotalQuantity = 1,
                            Level = 1,
                            ComponentIndex = componentIndex++,
                            AssemblyLevel = 1
                        };
                        
                        if (!string.IsNullOrEmpty(bomItem.ChildPartNumber) || !string.IsNullOrEmpty(bomItem.ComponentType))
                        {
                            xmlData.BomItems.Add(bomItem);
                        }
                    }
                }
            }
            
            Console.WriteLine($"DEBUG: Total BOM items extracted for {xmlData.PartNumber}: {xmlData.BomItems.Count}");
        }
        
        static string GetXmlAttribute(XmlNode parentNode, string attributeName)
        {
            var attrNode = parentNode.SelectSingleNode($"attribute[@name='{attributeName}']");
            return attrNode?.Attributes?["value"]?.Value ?? "";
        }
        
        static decimal? ParseDecimal(string value)
        {
            return decimal.TryParse(value, out var result) ? result : null;
        }
        
        static int? ParseInt(string value)
        {
            return int.TryParse(value, out var result) ? result : null;
        }
        
        static void InsertXmlData(SqliteConnection connection, XmlPartData xmlData)
        {
            using var transaction = connection.BeginTransaction();
            try
            {
                // First, delete existing records for this file if they exist
                var deleteCmd = new SqliteCommand(@"
                    DELETE FROM Components WHERE XMLFileID IN (SELECT ID FROM XMLFiles WHERE FilePath = @FilePath);
                    DELETE FROM ManufacturingFlags WHERE XMLFileID IN (SELECT ID FROM XMLFiles WHERE FilePath = @FilePath);
                    DELETE FROM PartData WHERE XMLFileID IN (SELECT ID FROM XMLFiles WHERE FilePath = @FilePath);
                    DELETE FROM XMLFiles WHERE FilePath = @FilePath", 
                    connection, transaction);
                deleteCmd.Parameters.AddWithValue("@FilePath", xmlData.FilePath);
                deleteCmd.ExecuteNonQuery();
                
                // Insert updated XML file record
                var xmlFileCmd = new SqliteCommand(@"
                    INSERT INTO XMLFiles (FilePath, FileName, PartNumber, Revision, Release, FileModifiedDate, ParsedDate)
                    VALUES (@FilePath, @FileName, @PartNumber, @Revision, @Release, @FileModifiedDate, @ParsedDate);
                    SELECT last_insert_rowid();", 
                    connection, transaction);
                
                xmlFileCmd.Parameters.AddWithValue("@FilePath", xmlData.FilePath);
                xmlFileCmd.Parameters.AddWithValue("@FileName", xmlData.FileName);
                xmlFileCmd.Parameters.AddWithValue("@PartNumber", xmlData.PartNumber);
                xmlFileCmd.Parameters.AddWithValue("@Revision", xmlData.Revision);
                xmlFileCmd.Parameters.AddWithValue("@Release", xmlData.Release);
                xmlFileCmd.Parameters.AddWithValue("@FileModifiedDate", xmlData.FileModifiedDate.ToString("yyyy-MM-dd HH:mm:ss"));
                xmlFileCmd.Parameters.AddWithValue("@ParsedDate", xmlData.ParsedDate.ToString("yyyy-MM-dd HH:mm:ss"));
                
                var xmlFileId = Convert.ToInt32(xmlFileCmd.ExecuteScalar());
                
                // Insert part data
                var partDataCmd = new SqliteCommand(@"
                    INSERT INTO PartData (XMLFileID, PartNumber, Revision, Release, Description, MakeBuy, Material, 
                                         Thickness, Weight, MaxX, MaxY, MaxZ, Rotation, GangQty, RawMaterialNumber)
                    VALUES (@XMLFileID, @PartNumber, @Revision, @Release, @Description, @MakeBuy, @Material,
                           @Thickness, @Weight, @MaxX, @MaxY, @MaxZ, @Rotation, @GangQty, @RawMaterialNumber)",
                    connection, transaction);
                
                partDataCmd.Parameters.AddWithValue("@XMLFileID", xmlFileId);
                partDataCmd.Parameters.AddWithValue("@PartNumber", xmlData.PartNumber);
                partDataCmd.Parameters.AddWithValue("@Revision", xmlData.Revision);
                partDataCmd.Parameters.AddWithValue("@Release", xmlData.Release);
                partDataCmd.Parameters.AddWithValue("@Description", xmlData.Description ?? "");
                partDataCmd.Parameters.AddWithValue("@MakeBuy", xmlData.MakeBuy ?? "");
                partDataCmd.Parameters.AddWithValue("@Material", xmlData.Material ?? "");
                partDataCmd.Parameters.AddWithValue("@Thickness", (object?)xmlData.Thickness ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@Weight", (object?)xmlData.Weight ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@MaxX", (object?)xmlData.MaxX ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@MaxY", (object?)xmlData.MaxY ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@MaxZ", (object?)xmlData.MaxZ ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@Rotation", (object?)xmlData.Rotation ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@GangQty", (object?)xmlData.GangQty ?? DBNull.Value);
                partDataCmd.Parameters.AddWithValue("@RawMaterialNumber", xmlData.RawMaterialNumber ?? "");
                
                partDataCmd.ExecuteNonQuery();
                
                // Insert manufacturing flags
                var flagsCmd = new SqliteCommand(@"
                    INSERT INTO ManufacturingFlags (XMLFileID, PartNumber, Laser, Punch, Saw, Shear, Powder,
                                                   LoosePart, ShipLoose, AssemblyCut, HardwareLot, Template, TemplateCut)
                    VALUES (@XMLFileID, @PartNumber, @Laser, @Punch, @Saw, @Shear, @Powder,
                           @LoosePart, @ShipLoose, @AssemblyCut, @HardwareLot, @Template, @TemplateCut)",
                    connection, transaction);
                
                flagsCmd.Parameters.AddWithValue("@XMLFileID", xmlFileId);
                flagsCmd.Parameters.AddWithValue("@PartNumber", xmlData.PartNumber);
                flagsCmd.Parameters.AddWithValue("@Laser", xmlData.Laser ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@Punch", xmlData.Punch ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@Saw", xmlData.Saw ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@Shear", xmlData.Shear ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@Powder", xmlData.Powder ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@LoosePart", xmlData.LoosePart ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@ShipLoose", xmlData.ShipLoose ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@AssemblyCut", xmlData.AssemblyCut ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@HardwareLot", xmlData.HardwareLot ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@Template", xmlData.Template ? 1 : 0);
                flagsCmd.Parameters.AddWithValue("@TemplateCut", xmlData.TemplateCut ? 1 : 0);
                
                flagsCmd.ExecuteNonQuery();
                
                // Insert component/BOM items
                foreach (var bomItem in xmlData.BomItems)
                {
                    var componentCmd = new SqliteCommand(@"
                        INSERT INTO Components (XMLFileID, ParentPartNumber, ComponentPartNumber, ComponentDescription,
                                              ComponentType, Quantity, TotalQuantity, Material, Thickness, ComponentIndex, AssemblyLevel,
                                              Weight, MaxX, MaxY, MaxZ, Rotation, GangQty, RawMaterialNumber)
                        VALUES (@XMLFileID, @ParentPartNumber, @ComponentPartNumber, @ComponentDescription,
                               @ComponentType, @Quantity, @TotalQuantity, @Material, @Thickness, @ComponentIndex, @AssemblyLevel,
                               @Weight, @MaxX, @MaxY, @MaxZ, @Rotation, @GangQty, @RawMaterialNumber)",
                        connection, transaction);
                    
                    componentCmd.Parameters.AddWithValue("@XMLFileID", xmlFileId);
                    componentCmd.Parameters.AddWithValue("@ParentPartNumber", xmlData.PartNumber);
                    componentCmd.Parameters.AddWithValue("@ComponentPartNumber", bomItem.ChildPartNumber ?? "");
                    componentCmd.Parameters.AddWithValue("@ComponentDescription", bomItem.Description ?? "");
                    componentCmd.Parameters.AddWithValue("@ComponentType", bomItem.ComponentType ?? "");
                    componentCmd.Parameters.AddWithValue("@Quantity", bomItem.Quantity);
                    componentCmd.Parameters.AddWithValue("@TotalQuantity", bomItem.TotalQuantity);
                    componentCmd.Parameters.AddWithValue("@Material", bomItem.Material ?? "");
                    componentCmd.Parameters.AddWithValue("@Thickness", (object?)bomItem.Thickness ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@ComponentIndex", bomItem.ComponentIndex);
                    componentCmd.Parameters.AddWithValue("@AssemblyLevel", bomItem.AssemblyLevel);
                    // Enhanced PartData properties
                    componentCmd.Parameters.AddWithValue("@Weight", (object?)bomItem.Weight ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@MaxX", (object?)bomItem.MaxX ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@MaxY", (object?)bomItem.MaxY ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@MaxZ", (object?)bomItem.MaxZ ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@Rotation", (object?)bomItem.Rotation ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@GangQty", (object?)bomItem.GangQty ?? DBNull.Value);
                    componentCmd.Parameters.AddWithValue("@RawMaterialNumber", bomItem.RawMaterialNumber ?? "");
                    
                    componentCmd.ExecuteNonQuery();
                }
                
                transaction.Commit();
                
                // Calculate total quantities considering assembly hierarchy
                CalculateTotalQuantities(connection, xmlFileId, xmlData.PartNumber);
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
        
        static void CalculateTotalQuantities(SqliteConnection connection, int xmlFileId, string parentPartNumber)
        {
            // This method will calculate total quantities considering parent assembly multipliers
            // For now, we'll implement a simple version that sets TotalQuantity = Quantity
            // In a future enhancement, this would traverse the assembly hierarchy
            
            var updateCmd = new SqliteCommand(@"
                UPDATE Components 
                SET TotalQuantity = Quantity 
                WHERE XMLFileID = @XMLFileID", connection);
            
            updateCmd.Parameters.AddWithValue("@XMLFileID", xmlFileId);
            updateCmd.ExecuteNonQuery();
            
            // Future enhancement: Recursive calculation of quantities through assembly hierarchy
            // This would require identifying parent-child relationships and multiplying quantities
        }
        
        static void ShowProcessingSummary()
        {
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            var cmd = new SqliteCommand(@"
                SELECT 
                    COUNT(*) as TotalFiles,
                    COUNT(DISTINCT xf.PartNumber) as UniqueParts,
                    COUNT(DISTINCT pd.Material) as UniqueMaterials
                FROM XMLFiles xf
                LEFT JOIN PartData pd ON xf.ID = pd.XMLFileID", connection);
            
            using var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                Console.WriteLine();
                Console.WriteLine("=== PROCESSING SUMMARY ===");
                Console.WriteLine($"Total XML files processed: {reader["TotalFiles"]}");
                Console.WriteLine($"Unique part numbers: {reader["UniqueParts"]}");
                Console.WriteLine($"Unique materials: {reader["UniqueMaterials"]}");
                Console.WriteLine($"Database file: {Path.GetFullPath(DatabasePath)}");
            }
        }
        
        // Sample query function to demonstrate database capabilities
        static void ShowSampleQueries()
        {
            Console.WriteLine("=== SAMPLE DATABASE QUERIES ===");
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();

            // First check what tables exist
            Console.WriteLine("\nDatabase Tables:");
            using (var cmd = new SqliteCommand("SELECT name FROM sqlite_master WHERE type='table'", connection))
            {
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Console.WriteLine($"  - {reader["name"]}");
                }
            }

            // Query 1: Sample parts with materials
            Console.WriteLine("\n1. Sample Parts with Material Data:");
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    xf.PartNumber,
                    xf.Revision,
                    pd.Material,
                    pd.Thickness,
                    pd.Weight
                FROM XMLFiles xf
                LEFT JOIN PartData pd ON xf.ID = pd.XMLFileID
                WHERE pd.Material IS NOT NULL AND pd.Material != ''
                LIMIT 5", connection))
            {
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Console.WriteLine($"  {reader["PartNumber"]} Rev{reader["Revision"]} - {reader["Material"]} (Thickness: {reader["Thickness"]}, Weight: {reader["Weight"]})");
                }
            }

            // Query 2: Hardware parts
            Console.WriteLine("\n2. Hardware Parts (HDW prefix):");
            using (var cmd = new SqliteCommand(@"
                SELECT PartNumber, COUNT(*) as RevisionCount
                FROM XMLFiles 
                WHERE PartNumber LIKE 'HDW%'
                GROUP BY PartNumber
                ORDER BY RevisionCount DESC
                LIMIT 5", connection))
            {
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Console.WriteLine($"  {reader["PartNumber"]} - {reader["RevisionCount"]} revisions");
                }
            }

            // Query 3: Check if BOM table exists before querying
            bool bomTableExists = false;
            using (var cmd = new SqliteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name='BomItems'", connection))
            {
                using var reader = cmd.ExecuteReader();
                bomTableExists = reader.Read();
            }

            if (bomTableExists)
            {
                Console.WriteLine("\n3. BOM Relationships:");
                using (var cmd = new SqliteCommand(@"
                    SELECT 
                        xf.PartNumber as Parent,
                        bi.ChildPartNumber as Child,
                        bi.Quantity
                    FROM XMLFiles xf
                    JOIN BomItems bi ON xf.ID = bi.XMLFileID
                    WHERE bi.ChildPartNumber IS NOT NULL AND bi.ChildPartNumber != ''
                    LIMIT 5", connection))
                {
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        Console.WriteLine($"  {reader["Parent"]} uses {reader["Quantity"]}x {reader["Child"]}");
                    }
                }
            }
            else
            {
                Console.WriteLine("\n3. BOM Relationships: No BOM data found in XMLs");
            }
            
            Console.WriteLine("\n🎯 Your XML intelligence database is ready for VBA integration!");
        }
        
        static void ShowHelp()
        {
            Console.WriteLine("XML Index Builder - Keep your database synchronized with new XML files");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  XMLIndexer.exe [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --full, -f         Force full scan (reprocess all files)");
            Console.WriteLine("  --incremental, -i  Incremental only (skip test, process new files only)");
            Console.WriteLine("  --explore, -e      Explore database content (no processing)");
            Console.WriteLine("  --search, -s       Search for specific part/file (use: --search PARTNUMBER)");
            Console.WriteLine("  --help, -h         Show this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  XMLIndexer.exe                # Smart update (default)");
            Console.WriteLine("  XMLIndexer.exe --incremental  # Process only new/modified files");
            Console.WriteLine("  XMLIndexer.exe --full         # Reprocess everything");
            Console.WriteLine("  XMLIndexer.exe --explore      # View database content");
            Console.WriteLine("  XMLIndexer.exe --search TUBE  # Find all parts with 'TUBE' in name");
        }
        
        static List<string> GetFilesToProcess(List<string> allFiles, bool forceFullScan)
        {
            if (forceFullScan)
            {
                Console.WriteLine("Force full scan requested - will process all files");
                return allFiles;
            }
            
            var filesToProcess = new List<string>();
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            
            foreach (var filePath in allFiles)
            {
                var fileInfo = new FileInfo(filePath);
                
                // Check if file exists in database and get its stored modification date
                using var cmd = new SqliteCommand(@"
                    SELECT FileModifiedDate 
                    FROM XMLFiles 
                    WHERE FilePath = @filePath", connection);
                cmd.Parameters.AddWithValue("@filePath", filePath);
                
                var storedModDate = cmd.ExecuteScalar();
                
                if (storedModDate == null)
                {
                    // File not in database - needs processing
                    filesToProcess.Add(filePath);
                }
                else if (storedModDate is DateTime storedDate)
                {
                    // File exists - check if it's been modified
                    if (fileInfo.LastWriteTime > storedDate.AddSeconds(1)) // 1 second tolerance
                    {
                        filesToProcess.Add(filePath);
                    }
                }
            }
            
            return filesToProcess;
        }
        
        static void ExploreDatabase()
        {
            Console.WriteLine("🔍 EXPLORING XML DATABASE CONTENT");
            Console.WriteLine($"Database: {Path.GetFullPath(DatabasePath)}");
            Console.WriteLine();
            
            if (!File.Exists(DatabasePath))
            {
                Console.WriteLine("❌ Database file not found. Run indexer first to create database.");
                return;
            }
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();

            // 1. Database overview
            Console.WriteLine("📊 DATABASE OVERVIEW:");
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    COUNT(*) as TotalFiles,
                    COUNT(DISTINCT PartNumber) as UniqueParts,
                    MIN(ParsedDate) as FirstProcessed,
                    MAX(ParsedDate) as LastProcessed
                FROM XMLFiles", connection))
            {
                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    Console.WriteLine($"   Total Files: {reader["TotalFiles"]}");
                    Console.WriteLine($"   Unique Parts: {reader["UniqueParts"]}");
                    Console.WriteLine($"   First Processed: {reader["FirstProcessed"]}");
                    Console.WriteLine($"   Last Processed: {reader["LastProcessed"]}");
                }
            }

            Console.WriteLine();

            // 2. Sample file paths and part numbers
            Console.WriteLine("📁 SAMPLE FILES (Recent 10):");
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    PartNumber,
                    Revision,
                    Release,
                    FileName
                FROM XMLFiles 
                ORDER BY ParsedDate DESC
                LIMIT 10", connection))
            {
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Console.WriteLine($"   {reader["PartNumber"]} Rev{reader["Revision"]} Rel{reader["Release"]}");
                    Console.WriteLine($"     File: {reader["FileName"]}");
                    Console.WriteLine();
                }
            }

            // 3. Part data samples with materials
            Console.WriteLine("🔧 PART DATA WITH MATERIALS:");
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    pd.PartNumber,
                    pd.Revision,
                    pd.Material,
                    pd.Thickness,
                    pd.Weight,
                    pd.Description
                FROM PartData pd
                WHERE pd.Material IS NOT NULL AND pd.Material != '' AND pd.Material != 'NULL'
                LIMIT 10", connection))
            {
                using var reader = cmd.ExecuteReader();
                int count = 0;
                while (reader.Read())
                {
                    count++;
                    Console.WriteLine($"   Part: {reader["PartNumber"]} Rev{reader["Revision"]}");
                    Console.WriteLine($"     Material: {reader["Material"]}");
                    Console.WriteLine($"     Thickness: {reader["Thickness"]}");
                    Console.WriteLine($"     Weight: {reader["Weight"]}");
                    Console.WriteLine($"     Description: {reader["Description"]}");
                    Console.WriteLine();
                }
                if (count == 0)
                {
                    Console.WriteLine("   No material data found in current records");
                }
            }

            // 4. Part number patterns
            Console.WriteLine("🏷️ PART NUMBER PATTERNS:");
            string[] prefixes = { "HDW", "STR", "BRK", "PLT", "TUB", "ANG", "FLT", "WLD", "ASS", "SHT" };

            foreach (var prefix in prefixes)
            {
                using (var cmd = new SqliteCommand($@"
                    SELECT COUNT(*) as Count
                    FROM XMLFiles 
                    WHERE PartNumber LIKE '{prefix}%'", connection))
                {
                    var count = cmd.ExecuteScalar();
                    if (Convert.ToInt32(count) > 0)
                    {
                        Console.WriteLine($"   {prefix}* pattern: {count} parts");
                    }
                }
            }

            Console.WriteLine();

            // 5. Sample actual part numbers
            Console.WriteLine("🔢 SAMPLE PART NUMBERS BY TYPE:");
            foreach (var prefix in new[] { "HDW", "STR", "PLT" })
            {
                using (var cmd = new SqliteCommand($@"
                    SELECT PartNumber, Revision
                    FROM XMLFiles 
                    WHERE PartNumber LIKE '{prefix}%'
                    ORDER BY PartNumber
                    LIMIT 3", connection))
                {
                    using var reader = cmd.ExecuteReader();
                    Console.WriteLine($"   {prefix} Examples:");
                    while (reader.Read())
                    {
                        Console.WriteLine($"     {reader["PartNumber"]} Rev{reader["Revision"]}");
                    }
                }
            }

            Console.WriteLine();
            
            // 6. Let's see what data we DO have in PartData
            Console.WriteLine("🔧 PARTDATA TABLE ANALYSIS:");
            using (var cmd = new SqliteCommand(@"
                SELECT COUNT(*) as Total FROM PartData", connection))
            {
                var total = cmd.ExecuteScalar();
                Console.WriteLine($"   Total PartData records: {total}");
            }
            
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    COUNT(CASE WHEN Material IS NOT NULL AND Material != '' THEN 1 END) as WithMaterial,
                    COUNT(CASE WHEN Description IS NOT NULL AND Description != '' THEN 1 END) as WithDescription,
                    COUNT(CASE WHEN Thickness IS NOT NULL THEN 1 END) as WithThickness
                FROM PartData", connection))
            {
                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    Console.WriteLine($"   Records with Material: {reader["WithMaterial"]}");
                    Console.WriteLine($"   Records with Description: {reader["WithDescription"]}");
                    Console.WriteLine($"   Records with Thickness: {reader["WithThickness"]}");
                }
            }
            
            // Show ANY PartData sample
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    PartNumber,
                    Material,
                    Description,
                    Thickness,
                    Weight
                FROM PartData 
                LIMIT 5", connection))
            {
                using var reader = cmd.ExecuteReader();
                Console.WriteLine($"   Sample PartData records (any data):");
                while (reader.Read())
                {
                    Console.WriteLine($"     {reader["PartNumber"]} - Mat:'{reader["Material"]}' Desc:'{reader["Description"]}' T:{reader["Thickness"]}");
                }
            }

            Console.WriteLine();
            Console.WriteLine("🎯 Database exploration complete!");
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
        
        static void SearchDatabase(string[] args)
        {
            // Get search term from arguments
            string searchTerm = "";
            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "--search" || args[i] == "-s") && i + 1 < args.Length)
                {
                    searchTerm = args[i + 1];
                    break;
                }
            }
            
            if (string.IsNullOrEmpty(searchTerm))
            {
                Console.WriteLine("❌ Please provide a search term.");
                Console.WriteLine("Usage: XMLIndexer --search PARTNUMBER");
                Console.WriteLine("Example: XMLIndexer --search TUBE");
                return;
            }
            
            Console.WriteLine($"🔍 SEARCHING XML DATABASE FOR: '{searchTerm}'");
            Console.WriteLine($"Database: {Path.GetFullPath(DatabasePath)}");
            Console.WriteLine();
            
            if (!File.Exists(DatabasePath))
            {
                Console.WriteLine("❌ Database file not found. Run indexer first to create database.");
                return;
            }
            
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();

            // Search XMLFiles for matching parts
            Console.WriteLine("📁 MATCHING XML FILES:");
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    xf.ID,
                    xf.PartNumber,
                    xf.Revision,
                    xf.Release,
                    xf.FileName,
                    xf.FilePath,
                    xf.FileModifiedDate,
                    xf.ParsedDate
                FROM XMLFiles xf
                WHERE xf.PartNumber LIKE @searchTerm 
                   OR xf.FileName LIKE @searchTerm
                ORDER BY xf.PartNumber, xf.Revision", connection))
            {
                cmd.Parameters.AddWithValue("@searchTerm", $"%{searchTerm}%");
                using var reader = cmd.ExecuteReader();
                
                int count = 0;
                while (reader.Read())
                {
                    count++;
                    Console.WriteLine($"   [{count}] {reader["PartNumber"]} Rev{reader["Revision"]} Rel{reader["Release"]}");
                    Console.WriteLine($"       File: {reader["FileName"]}");
                    Console.WriteLine($"       Path: {reader["FilePath"]}");
                    Console.WriteLine($"       Modified: {reader["FileModifiedDate"]}");
                    Console.WriteLine($"       Processed: {reader["ParsedDate"]}");
                    Console.WriteLine($"       Database ID: {reader["ID"]}");
                    Console.WriteLine();
                }
                
                if (count == 0)
                {
                    Console.WriteLine($"   No files found matching '{searchTerm}'");
                    Console.WriteLine();
                    Console.WriteLine("💡 Try searching for:");
                    Console.WriteLine("   - Part number: TUBE, STRUT, UPPER, etc.");
                    Console.WriteLine("   - Part of filename: WELDMENT, BARRIER, etc.");
                    return;
                }
                
                Console.WriteLine($"Found {count} matching files.");
                Console.WriteLine();
                
                // Ask if user wants detailed data for a specific file
                Console.Write("Enter file number [1-" + count + "] to see detailed data (or press Enter to exit): ");
                var input = Console.ReadLine();
                
                if (int.TryParse(input, out int fileNumber) && fileNumber >= 1 && fileNumber <= count)
                {
                    ShowDetailedFileData(connection, searchTerm, fileNumber);
                }
            }
        }
        
        static void ShowDetailedFileData(SqliteConnection connection, string searchTerm, int fileNumber)
        {
            Console.WriteLine();
            Console.WriteLine("📋 DETAILED XML DATA:");
            
            // Get the specific file ID for the selected file
            using (var cmd = new SqliteCommand(@"
                SELECT 
                    xf.ID,
                    xf.PartNumber,
                    xf.Revision,
                    xf.Release,
                    xf.FileName,
                    xf.FilePath
                FROM XMLFiles xf
                WHERE xf.PartNumber LIKE @searchTerm 
                   OR xf.FileName LIKE @searchTerm
                ORDER BY xf.PartNumber, xf.Revision
                LIMIT 1 OFFSET @offset", connection))
            {
                cmd.Parameters.AddWithValue("@searchTerm", $"%{searchTerm}%");
                cmd.Parameters.AddWithValue("@offset", fileNumber - 1);
                
                using var reader = cmd.ExecuteReader();
                if (!reader.Read())
                {
                    Console.WriteLine("File not found.");
                    return;
                }
                
                var fileId = reader["ID"];
                var partNumber = reader["PartNumber"];
                var revision = reader["Revision"];
                var fileName = reader["FileName"];
                var filePath = reader["FilePath"];
                
                Console.WriteLine($"Part: {partNumber} Rev{revision}");
                Console.WriteLine($"File: {fileName}");
                Console.WriteLine($"Path: {filePath}");
                Console.WriteLine();
                
                reader.Close();
                
                // Get PartData for this file
                Console.WriteLine("🔧 PART DATA EXTRACTED:");
                using (var partCmd = new SqliteCommand(@"
                    SELECT 
                        PartNumber,
                        Revision,
                        Material,
                        Thickness,
                        Weight,
                        MaxX,
                        MaxY,
                        MaxZ,
                        Description,
                        Finish,
                        Notes
                    FROM PartData 
                    WHERE XMLFileID = @fileId", connection))
                {
                    partCmd.Parameters.AddWithValue("@fileId", fileId);
                    using var partReader = partCmd.ExecuteReader();
                    
                    if (partReader.Read())
                    {
                        Console.WriteLine($"   Part Number: {partReader["PartNumber"]}");
                        Console.WriteLine($"   Revision: {partReader["Revision"]}");
                        Console.WriteLine($"   Material: '{partReader["Material"]}'");
                        Console.WriteLine($"   Thickness: {partReader["Thickness"]}");
                        Console.WriteLine($"   Weight: {partReader["Weight"]}");
                        Console.WriteLine($"   Dimensions: {partReader["MaxX"]} x {partReader["MaxY"]} x {partReader["MaxZ"]}");
                        Console.WriteLine($"   Description: '{partReader["Description"]}'");
                        Console.WriteLine($"   Finish: '{partReader["Finish"]}'");
                        Console.WriteLine($"   Notes: '{partReader["Notes"]}'");
                    }
                    else
                    {
                        Console.WriteLine("   No detailed part data found for this file.");
                    }
                }
                
                Console.WriteLine();
                
                // Get ManufacturingFlags for this file
                Console.WriteLine("🏭 MANUFACTURING FLAGS:");
                using (var flagCmd = new SqliteCommand(@"
                    SELECT *
                    FROM ManufacturingFlags 
                    WHERE XMLFileID = @fileId", connection))
                {
                    flagCmd.Parameters.AddWithValue("@fileId", fileId);
                    using var flagReader = flagCmd.ExecuteReader();
                    
                    if (flagReader.Read())
                    {
                        // Display all manufacturing flag columns
                        for (int i = 0; i < flagReader.FieldCount; i++)
                        {
                            if (flagReader.GetName(i) != "ID" && flagReader.GetName(i) != "XMLFileID")
                            {
                                Console.WriteLine($"   {flagReader.GetName(i)}: {flagReader[i]}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("   No manufacturing flags found for this file.");
                    }
                }
            }
            
            Console.WriteLine();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
        
        static void ApplyMrpSchema()
        {
            Console.WriteLine("=== APPLYING MRP SCHEMA ===");
            
            try
            {
                string schemaPath = "MRP_Schema.sql";
                if (!File.Exists(schemaPath))
                {
                    Console.WriteLine($"❌ Schema file not found: {schemaPath}");
                    return;
                }
                
                string schemaSql = File.ReadAllText(schemaPath);
                Console.WriteLine($"✅ Read schema file: {schemaPath}");
                
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                using var command = new SqliteCommand(schemaSql, connection);
                command.ExecuteNonQuery();
                
                Console.WriteLine("✅ MRP schema applied successfully");
                
                // Test that the table was created
                using var testCommand = new SqliteCommand(
                    "SELECT COUNT(*) FROM MrpPriorityList", connection);
                var count = Convert.ToInt32(testCommand.ExecuteScalar());
                Console.WriteLine($"✅ MrpPriorityList table has {count} records");
                
                // Check for test job
                using var testJobCommand = new SqliteCommand(
                    "SELECT COUNT(*) FROM MrpPriorityList WHERE JobNumber = 'H1319-0000'", connection);
                var testJobCount = Convert.ToInt32(testJobCommand.ExecuteScalar());
                Console.WriteLine($"✅ Test job (H1319-0000) found: {testJobCount} times");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error applying schema: {ex.Message}");
            }
        }
    }
    
    public class XmlPartData
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public string PartNumber { get; set; } = "";
        public string Revision { get; set; } = "";
        public string Release { get; set; } = "";
        public string Description { get; set; } = "";
        public string MakeBuy { get; set; } = "";
        public string Material { get; set; } = "";
        public string RawMaterialNumber { get; set; } = "";
        public decimal? Thickness { get; set; }
        public decimal? Weight { get; set; }
        public decimal? MaxX { get; set; }
        public decimal? MaxY { get; set; }
        public decimal? MaxZ { get; set; }
        public int? Rotation { get; set; }
        public int? GangQty { get; set; }
        public DateTime ParsedDate { get; set; }
        public DateTime FileModifiedDate { get; set; }
        
        // Manufacturing flags
        public bool Laser { get; set; }
        public bool Punch { get; set; }
        public bool Saw { get; set; }
        public bool Shear { get; set; }
        public bool Powder { get; set; }
        public bool LoosePart { get; set; }
        public bool ShipLoose { get; set; }
        public bool AssemblyCut { get; set; }
        public bool HardwareLot { get; set; }
        public bool Template { get; set; }
        public bool TemplateCut { get; set; }
        
        public List<BomItem> BomItems { get; set; } = new();
    }
    
    public class BomItem
    {
        public string ChildPartNumber { get; set; } = "";
        public string ChildRevision { get; set; } = "";
        public decimal Quantity { get; set; }
        public decimal TotalQuantity { get; set; } // Calculated quantity including parent multipliers
        public int Level { get; set; }
        public string ComponentType { get; set; } = ""; // make, buy, stock, etc.
        public string Description { get; set; } = "";
        public string Material { get; set; } = "";
        public decimal? Thickness { get; set; }
        public int ComponentIndex { get; set; } = 0;
        public int AssemblyLevel { get; set; } = 1;
        
        // Additional PartData properties for components
        public decimal? Weight { get; set; }
        public decimal? MaxX { get; set; }
        public decimal? MaxY { get; set; }
        public decimal? MaxZ { get; set; }
        public int? Rotation { get; set; }
        public int? GangQty { get; set; }
        public string RawMaterialNumber { get; set; } = "";
    }
}
