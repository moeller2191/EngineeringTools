using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.Sqlite;
using System.IO;

namespace XMLIndexer
{
    public class MrpDataManager
    {
        private readonly string _connectionString;
        
        public MrpDataManager(string databasePath)
        {
            _connectionString = $"Data Source={databasePath}";
            InitializeDatabase();
        }
        
        private void InitializeDatabase()
        {
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();
            
            // Create MrpPriorityList table if it doesn't exist
            var createTableCommand = new SqliteCommand(@"
                CREATE TABLE IF NOT EXISTS MrpPriorityList (
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
                    
                    UNIQUE(JobNumber, PartNumber, Revision)
                )", connection);
            createTableCommand.ExecuteNonQuery();
            
            // Create indexes for performance
            var createIndexCommands = new[]
            {
                "CREATE INDEX IF NOT EXISTS idx_mrp_job_number ON MrpPriorityList(JobNumber)",
                "CREATE INDEX IF NOT EXISTS idx_mrp_part_number ON MrpPriorityList(PartNumber)",
                "CREATE INDEX IF NOT EXISTS idx_mrp_priority ON MrpPriorityList(Priority)",
                "CREATE INDEX IF NOT EXISTS idx_mrp_status ON MrpPriorityList(Status)"
            };
            
            foreach (var sql in createIndexCommands)
            {
                var indexCommand = new SqliteCommand(sql, connection);
                indexCommand.ExecuteNonQuery();
            }
        }
        
        public class MrpItem
        {
            public int ID { get; set; }
            public string JobNumber { get; set; } = "";
            public string PartNumber { get; set; } = "";
            public string Revision { get; set; } = "";
            public int Quantity { get; set; } = 1;
            public string Description { get; set; } = "";
            public int Priority { get; set; }
            public string DueDate { get; set; } = "";
            public string Status { get; set; } = "";
            public string Customer { get; set; } = "";
            public string Program { get; set; } = "";
            public string Notes { get; set; } = "";
            public DateTime LastUpdated { get; set; }
            public string XmlStatus { get; set; } = "";
            public int HighestRelease { get; set; }
            public int ComponentCount { get; set; }
        }
        
        /// <summary>
        /// Get MRP data for a specific job number
        /// </summary>
        public List<MrpItem> GetMrpDataForJob(string jobNumber)
        {
            var items = new List<MrpItem>();
            
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();
            
            using var command = new SqliteCommand(@"
                SELECT m.*,
                       'No XML' as XmlStatus,
                       0 as HighestRelease,
                       0 as ComponentCount
                FROM MrpPriorityList m
                WHERE UPPER(m.JobNumber) LIKE UPPER(@jobNumber) 
                   OR UPPER(m.JobNumber) = UPPER(@exactJobNumber)
                ORDER BY m.Priority ASC", connection);
                
            command.Parameters.AddWithValue("@jobNumber", $"%{jobNumber}%");
            command.Parameters.AddWithValue("@exactJobNumber", jobNumber);
            
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                items.Add(new MrpItem
                {
                    ID = reader.GetInt32("ID"),
                    JobNumber = reader.GetString("JobNumber"),
                    PartNumber = reader.IsDBNull("PartNumber") ? "" : reader.GetString("PartNumber"),
                    Revision = reader.IsDBNull("Revision") ? "" : reader.GetString("Revision"),
                    Quantity = reader.IsDBNull("Quantity") ? 1 : reader.GetInt32("Quantity"),
                    Description = reader.IsDBNull("Description") ? "" : reader.GetString("Description"),
                    Priority = reader.IsDBNull("Priority") ? 1 : reader.GetInt32("Priority"),
                    DueDate = reader.IsDBNull("DueDate") ? "" : reader.GetString("DueDate"),
                    Status = reader.IsDBNull("Status") ? "" : reader.GetString("Status"),
                    Customer = reader.IsDBNull("Customer") ? "" : reader.GetString("Customer"),
                    Program = reader.IsDBNull("Program") ? "" : reader.GetString("Program"),
                    Notes = reader.IsDBNull("Notes") ? "" : reader.GetString("Notes"),
                    LastUpdated = reader.IsDBNull("LastUpdated") ? DateTime.Now : reader.GetDateTime("LastUpdated"),
                    XmlStatus = reader.IsDBNull("XmlStatus") ? "No XML" : reader.GetString("XmlStatus"),
                    HighestRelease = reader.IsDBNull("HighestRelease") ? 0 : reader.GetInt32("HighestRelease"),
                    ComponentCount = reader.IsDBNull("ComponentCount") ? 0 : reader.GetInt32("ComponentCount")
                });
            }
            
            return items;
        }
        
        /// <summary>
        /// Get all active MRP jobs
        /// </summary>
        public List<MrpItem> GetActiveMrpJobs()
        {
            var items = new List<MrpItem>();
            
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();
            
            using var command = new SqliteCommand(@"
                SELECT * FROM vw_MrpWithXmlStatus 
                ORDER BY Priority ASC", connection);
                
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                items.Add(new MrpItem
                {
                    ID = reader.GetInt32("ID"),
                    JobNumber = reader.GetString("JobNumber"),
                    PartNumber = reader.GetString("PartNumber"),
                    Revision = reader.IsDBNull("Revision") ? "" : reader.GetString("Revision"),
                    Quantity = reader.GetInt32("Quantity"),
                    Description = reader.IsDBNull("Description") ? "" : reader.GetString("Description"),
                    Priority = reader.GetInt32("Priority"),
                    DueDate = reader.IsDBNull("DueDate") ? "" : reader.GetString("DueDate"),
                    Status = reader.IsDBNull("Status") ? "" : reader.GetString("Status"),
                    Customer = reader.IsDBNull("Customer") ? "" : reader.GetString("Customer"),
                    Program = reader.IsDBNull("Program") ? "" : reader.GetString("Program"),
                    Notes = reader.IsDBNull("Notes") ? "" : reader.GetString("Notes"),
                    LastUpdated = reader.GetDateTime("LastUpdated"),
                    XmlStatus = reader.IsDBNull("XmlStatus") ? "" : reader.GetString("XmlStatus"),
                    HighestRelease = reader.IsDBNull("HighestRelease") ? 0 : reader.GetInt32("HighestRelease"),
                    ComponentCount = reader.GetInt32("ComponentCount")
                });
            }
            
            return items;
        }
        
        /// <summary>
        /// Update MRP data from external source (e.g., Excel file)
        /// </summary>
        public void UpdateMrpData(List<MrpItem> newData)
        {
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();
            
            using var transaction = connection.BeginTransaction();
            try
            {
                // Clear existing data
                using var deleteCommand = new SqliteCommand(
                    "DELETE FROM MrpPriorityList", connection, transaction);
                deleteCommand.ExecuteNonQuery();
                
                // Insert new data
                foreach (var item in newData)
                {
                    using var insertCommand = new SqliteCommand(@"
                        INSERT OR REPLACE INTO MrpPriorityList (JobNumber, PartNumber, Revision, Quantity, Description, Priority, DueDate, Status, Customer, Program, Notes, LastUpdated)
                        VALUES (@jobNumber, @partNumber, @revision, @quantity, @description, @priority, @dueDate, @status, @customer, @program, @notes, @lastUpdated)", connection, transaction);
                    
                    insertCommand.Parameters.AddWithValue("@jobNumber", item.JobNumber);
                    insertCommand.Parameters.AddWithValue("@partNumber", item.PartNumber);
                    insertCommand.Parameters.AddWithValue("@revision", item.Revision);
                    insertCommand.Parameters.AddWithValue("@quantity", item.Quantity);
                    insertCommand.Parameters.AddWithValue("@description", item.Description);
                    insertCommand.Parameters.AddWithValue("@priority", item.Priority);
                    insertCommand.Parameters.AddWithValue("@dueDate", item.DueDate);
                    insertCommand.Parameters.AddWithValue("@status", item.Status);
                    insertCommand.Parameters.AddWithValue("@customer", item.Customer);
                    insertCommand.Parameters.AddWithValue("@program", item.Program);
                    insertCommand.Parameters.AddWithValue("@notes", item.Notes);
                    insertCommand.Parameters.AddWithValue("@lastUpdated", DateTime.Now);
                    
                    insertCommand.ExecuteNonQuery();
                }
                
                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
        
        /// <summary>
        /// Get jobs that are ready for burning (have XML files assigned in Excel OR are I-jobs that can auto-assign)
        /// </summary>
        public List<MrpItem> GetJobsReadyForBurning(string excelPath)
        {
            var items = new List<MrpItem>();
            
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                return items; // Return empty list if no Excel file
            }
            
            try
            {
                // Read existing job/XML assignments from Excel
                var jobXmlAssignments = ReadJobXmlAssignmentsFromExcel(excelPath);
                
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();
                
                using var command = new SqliteCommand(@"
                    SELECT * FROM vw_MrpWithXmlStatus 
                    ORDER BY Priority ASC", connection);
                    
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    string jobNumber = reader.GetString("JobNumber");
                    string partNumber = reader.IsDBNull("PartNumber") ? "" : reader.GetString("PartNumber");
                    bool isIJob = jobNumber.StartsWith("I", StringComparison.OrdinalIgnoreCase);
                    
                    bool hasExcelAssignment = jobXmlAssignments.ContainsKey(jobNumber);
                    
                    // Job is ready for burning if:
                    // 1. Has Excel assignment (regular workflow), OR
                    // 2. Is an I-job with a part number (can auto-assign highest release XML)
                    if (hasExcelAssignment || (isIJob && !string.IsNullOrEmpty(partNumber)))
                    {
                        items.Add(new MrpItem
                        {
                            ID = reader.GetInt32("ID"),
                            JobNumber = jobNumber,
                            PartNumber = partNumber,
                            Revision = reader.IsDBNull("Revision") ? "" : reader.GetString("Revision"),
                            Quantity = reader.GetInt32("Quantity"),
                            Description = reader.IsDBNull("Description") ? "" : reader.GetString("Description"),
                            Priority = reader.GetInt32("Priority"),
                            DueDate = reader.IsDBNull("DueDate") ? "" : reader.GetString("DueDate"),
                            Status = reader.IsDBNull("Status") ? "" : reader.GetString("Status"),
                            Customer = reader.IsDBNull("Customer") ? "" : reader.GetString("Customer"),
                            Program = reader.IsDBNull("Program") ? "" : reader.GetString("Program"),
                            Notes = reader.IsDBNull("Notes") ? "" : reader.GetString("Notes"),
                            LastUpdated = reader.GetDateTime("LastUpdated"),
                            XmlStatus = reader.IsDBNull("XmlStatus") ? "" : reader.GetString("XmlStatus"),
                            HighestRelease = reader.IsDBNull("HighestRelease") ? 0 : reader.GetInt32("HighestRelease"),
                            ComponentCount = reader.GetInt32("ComponentCount")
                        });
                    }
                }
            }
            catch
            {
                // If Excel reading fails, return empty list
            }
            
            return items;
        }
        
        /// <summary>
        /// Read job/XML assignments from Excel file
        /// </summary>
        private Dictionary<string, string> ReadJobXmlAssignmentsFromExcel(string excelPath)
        {
            var assignments = new Dictionary<string, string>();
            
            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return assignments;
                
                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                
                dynamic workbook = excel.Workbooks.Open(excelPath);
                dynamic worksheet = workbook.Worksheets["Priority List"];
                
                // Find JobNumber and XMLFile columns
                int headerRow = 1;
                int jobNumberCol = -1;
                int xmlFileCol = -1;
                int lastCol = worksheet.UsedRange.Columns.Count;
                
                for (int col = 1; col <= lastCol; col++)
                {
                    var headerValue = GetCellValue(worksheet, headerRow, col)?.ToString()?.Trim();
                    if (string.Equals(headerValue, "JobNumber", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(headerValue, "Job Number", StringComparison.OrdinalIgnoreCase))
                    {
                        jobNumberCol = col;
                    }
                    else if (string.Equals(headerValue, "XMLFile", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(headerValue, "XML File", StringComparison.OrdinalIgnoreCase))
                    {
                        xmlFileCol = col;
                    }
                }
                
                // Read job/XML assignments
                if (jobNumberCol > 0 && xmlFileCol > 0)
                {
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    
                    for (int row = headerRow + 1; row <= lastRow; row++)
                    {
                        var jobNumber = GetCellValue(worksheet, row, jobNumberCol)?.ToString()?.Trim();
                        var xmlFile = GetCellValue(worksheet, row, xmlFileCol)?.ToString()?.Trim();
                        
                        if (!string.IsNullOrEmpty(jobNumber) && !string.IsNullOrEmpty(xmlFile))
                        {
                            assignments[jobNumber] = xmlFile;
                        }
                    }
                }
                
                workbook.Close(false);
                excel.Quit();
                
                // Clean up COM objects
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
            catch
            {
                // If Excel reading fails, return empty dictionary
            }
            
            return assignments;
        }
        
        /// <summary>
        /// Get XML file path for a job number from Excel, with auto-assignment for I-jobs
        /// </summary>
        public string GetXmlFilePathForJob(string jobNumber, string excelPath)
        {
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return "";
                
            // First check if there's already an assignment in Excel
            var assignments = ReadJobXmlAssignmentsFromExcel(excelPath);
            if (assignments.ContainsKey(jobNumber))
            {
                return assignments[jobNumber];
            }
            
            // For I-jobs, auto-assign highest release XML and save to Excel
            if (jobNumber.StartsWith("I", StringComparison.OrdinalIgnoreCase))
            {
                string autoXmlFile = GetHighestReleaseXmlForJob(jobNumber);
                if (!string.IsNullOrEmpty(autoXmlFile))
                {
                    // Save the auto-assignment back to Excel for future reference
                    var newAssignments = new Dictionary<string, string> { { jobNumber, autoXmlFile } };
                    UpdateExcelWithJobXmlAssignments(excelPath, newAssignments);
                    return autoXmlFile;
                }
            }
            
            return "";
        }
        
        /// <summary>
        /// Get highest release XML file for a job number (for I-job auto-assignment)
        /// </summary>
        private string GetHighestReleaseXmlForJob(string jobNumber)
        {
            try
            {
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();
                
                // For I-jobs, find the highest release XML for the part number
                using var command = new SqliteCommand(@"
                    SELECT x.FileName 
                    FROM XMLIndex x
                    JOIN vw_MrpWithXmlStatus m ON x.PartNumber = m.PartNumber 
                    WHERE m.JobNumber = @jobNumber 
                      AND x.Release = (
                          SELECT MAX(x2.Release) 
                          FROM XMLIndex x2 
                          WHERE x2.PartNumber = x.PartNumber
                      )
                    LIMIT 1", connection);
                    
                command.Parameters.AddWithValue("@jobNumber", jobNumber);
                
                var result = command.ExecuteScalar();
                return result?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }
        
        /// <summary>
        /// Check if engineering is completed for a job (has XML file assigned)
        /// </summary>
        public bool IsEngineeringCompleted(string jobNumber, string excelPath)
        {
            return !string.IsNullOrEmpty(GetXmlFilePathForJob(jobNumber, excelPath));
        }

        /// <summary>
        /// Import MRP data from Excel file (placeholder - would need Excel reading library)
        /// </summary>
        public void ImportFromExcelFile(string filePath)
        {
            // TODO: Implement Excel reading using ClosedXML or similar
            // For now, this is a placeholder that shows the structure
            
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel file not found: {filePath}");
            }
            
            // Placeholder: In a real implementation, you would:
            // 1. Read the Excel file
            // 2. Parse the data into MrpItem objects
            // 3. Call UpdateMrpData() with the parsed data
            
            Console.WriteLine($"Would import MRP data from: {filePath}");
        }
        
        /// <summary>
        /// Import MRP data from Excel file
        /// </summary>
        public bool ImportFromExcel(string excelPath)
        {
            try
            {
                var newMrpItems = ReadExcelData(excelPath);
                if (newMrpItems.Count > 0)
                {
                    UpdateMrpData(newMrpItems);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Excel import failed: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Read MRP data from Excel file
        /// </summary>
        private List<MrpItem> ReadExcelData(string excelPath)
        {
            var items = new List<MrpItem>();
            try
            {
                dynamic excelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                var workbookObj = excelApp.Workbooks.Open(excelPath);
                
                // Find the "Priority List" worksheet
                var worksheetObj = workbookObj.Worksheets["Priority List"];
                var usedRangeObj = worksheetObj.UsedRange;
                int totalRows = usedRangeObj.Rows.Count;
                var processedJobs = new HashSet<string>(); // Track unique job numbers to avoid duplicates
                
                Console.WriteLine($"Processing {totalRows} rows from Priority List worksheet...");
                Console.WriteLine("Filtering for unique jobs (removing routing duplicates)...");

                for (int excelRow = 2; excelRow <= totalRows; excelRow++)
                {
                    // Show progress every 10000 rows
                    if (excelRow % 10000 == 0)
                    {
                        Console.WriteLine($"  Processed {excelRow}/{totalRows} rows, imported {items.Count} unique jobs so far...");
                    }
                    
                    // TRY DIFFERENT MAPPING - part number might be in column 2 or 1
                    // FINAL CORRECT MAPPING BASED ON HEADERS
                    var jobNum = GetCellValue(worksheetObj, excelRow, 3)?.ToString()?.Trim();    // fjobno
                    var qtyStr = GetCellValue(worksheetObj, excelRow, 4)?.ToString()?.Trim();    // fquantity
                    var partNum = GetCellValue(worksheetObj, excelRow, 5)?.ToString()?.Trim();   // fpartno
                    var desc = GetCellValue(worksheetObj, excelRow, 6)?.ToString()?.Trim();      // fdesc
                    var revision = GetCellValue(worksheetObj, excelRow, 23)?.ToString()?.Trim(); // fcudrev (NEW)
                    var stat = GetCellValue(worksheetObj, excelRow, 12)?.ToString()?.Trim();     // fstatus
                    var descMemo = GetCellValue(worksheetObj, excelRow, 22)?.ToString()?.Trim(); // fdescmemo
                    
                    // DEBUG: Print header row and first 10 data rows for columns 5 and 6
                    if (excelRow == 2) {
                        Console.WriteLine("=== HEADER ROW ===");
                        for (int col = 1; col <= 25; col++) {
                            var headerVal = GetCellValue(worksheetObj, 1, col)?.ToString()?.Trim();
                            Console.WriteLine($"Header Col {col}: '{headerVal ?? "NULL"}'");
                        }
                        Console.WriteLine("=== END HEADER ===");
                    }
                    if (excelRow <= 11) {
                        Console.WriteLine($"Row {excelRow}: Job='{jobNum}', Part='{partNum}', Rev='{revision}', Desc='{desc}'");
                    }

                    if (!string.IsNullOrEmpty(jobNum) && !string.IsNullOrEmpty(partNum))
                    {
                        // Skip if we've already processed this job number (removes routing duplicates)
                        string jobKey = $"{jobNum}-{partNum}";
                        if (processedJobs.Contains(jobKey))
                        {
                            continue; // Skip duplicate job entries from routing table
                        }
                        processedJobs.Add(jobKey);
                        
                        int qtyVal = 1;
                        int.TryParse(qtyStr, out qtyVal);
                        
                        // Use revision from fcudrev column, but fall back to memo extraction if NS
                        string extractedRevision = "";
                        string cleanNotes = descMemo ?? "";
                        
                        if (!string.IsNullOrEmpty(revision) && !revision.Equals("NS", StringComparison.OrdinalIgnoreCase))
                        {
                            // Use the revision field directly from fcudrev column
                            extractedRevision = revision;
                        }
                        else if (!string.IsNullOrEmpty(descMemo) && descMemo.StartsWith("`REV"))
                        {
                            // Fall back to extracting revision from memo field when fcudrev is NS
                            int revIndex = descMemo.IndexOf("`REV");
                            if (revIndex >= 0)
                            {
                                string revPart = descMemo.Substring(revIndex + 4); // Skip "`REV"
                                // Take first 3 characters for numeric part (000)
                                if (revPart.Length >= 3 && int.TryParse(revPart.Substring(0, 3), out int revNumber))
                                {
                                    extractedRevision = revNumber.ToString("00"); // Convert to two-digit format
                                }
                                
                                // Clean the revision prefix from notes (remove `REV000)
                                int endOfRev = revIndex + 7; // "`REV" + "000" = 7 characters
                                if (endOfRev < descMemo.Length)
                                {
                                    cleanNotes = descMemo.Substring(endOfRev).TrimStart(',', ' ');
                                }
                                else
                                {
                                    cleanNotes = "";
                                }
                            }
                        }
                        
                        var mrpItem = new MrpItem
                        {
                            JobNumber = jobNum,
                            PartNumber = partNum,
                            Revision = extractedRevision, // Use revision from fcudrev or extracted from memo if NS
                            Quantity = qtyVal,
                            Description = desc ?? "",
                            Priority = 1,
                            Status = stat ?? "",
                            Notes = cleanNotes, // Store cleaned memo when revision was extracted, otherwise full memo
                            LastUpdated = DateTime.Now
                        };
                        items.Add(mrpItem);
                    }
                }
                workbookObj.Close(false);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to read Excel file: {ex.Message}");
            }
            return items;
        }
        
        /// <summary>
        /// Update Excel with job/XML assignments
        /// </summary>
        public bool UpdateExcelWithJobXmlAssignments(string excelPath, Dictionary<string, string> assignments)
        {
            // TODO: Implement Excel writing for job/XML assignments
            // This is a placeholder for the Excel writing functionality
            return true; // Return true for now
        }

        /// <summary>
        /// Get list of all programmed parts from Press Programs table
        /// </summary>
        public List<string> GetProgrammedParts(string excelPath)
        {
            var programmedParts = new List<string>();
            
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return programmedParts;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return programmedParts;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Press Programs" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Press Programs", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                if (worksheet != null)
                {
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var formDetail = GetCellValue(worksheet, row, 1)?.ToString()?.Trim(); // Column A = Form Detail
                        if (!string.IsNullOrEmpty(formDetail))
                        {
                            programmedParts.Add(formDetail);
                        }
                    }
                }

                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return programmedParts;
            }
            catch
            {
                return programmedParts; // If Excel reading fails, return empty list
            }
        }

        /// <summary>
        /// Check multiple sales orders at once - optimized batch method
        /// Opens Excel once and checks all sales orders in a single session
        /// </summary>
        public Dictionary<string, bool> CheckSalesOrdersBatch(List<string> salesOrders, string excelPath)
        {
            var results = new Dictionary<string, bool>();
            
            // Initialize all as false
            foreach (var so in salesOrders)
            {
                results[so] = false;
            }
            
            if (salesOrders.Count == 0 || string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return results;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return results;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Checked Sales Orders" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Checked Sales Orders", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                if (worksheet != null)
                {
                    // Read all checked sales orders once
                    var checkedSalesOrders = new HashSet<string>();
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var cellValue = GetCellValue(worksheet, row, 2)?.ToString()?.Trim(); // Column B = SO
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            checkedSalesOrders.Add(cellValue.ToUpper());
                        }
                    }
                    
                    // Check each requested sales order against the set
                    foreach (var salesOrder in salesOrders)
                    {
                        string soNumber = salesOrder.Length >= 5 ? salesOrder.Substring(salesOrder.Length - 5) : salesOrder;
                        results[salesOrder] = checkedSalesOrders.Contains(soNumber.ToUpper());
                    }
                }

                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return results;
            }
            catch
            {
                return results; // If Excel reading fails, return all false
            }
        }

        /// <summary>
        /// Check if a sales order is already in the Checked Sales Orders table
        /// </summary>
        public bool CheckSalesOrder(string salesOrder, string excelPath)
        {
            if (string.IsNullOrEmpty(salesOrder) || string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return false;

            // Extract last 5 digits of sales order (matching original VBA logic)
            string soNumber = salesOrder.Length >= 5 ? salesOrder.Substring(salesOrder.Length - 5) : salesOrder;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return false;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Checked Sales Orders" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Checked Sales Orders", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                bool found = false;
                if (worksheet != null)
                {
                    // Find SO column (assuming column B based on CSV structure)
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var cellValue = GetCellValue(worksheet, row, 2)?.ToString()?.Trim(); // Column B = SO
                        if (string.Equals(cellValue, soNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            found = true;
                            break;
                        }
                    }
                }

                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return found;
            }
            catch
            {
                return false; // If Excel reading fails, assume not checked
            }
        }

        /// <summary>
        /// Add a sales order to the Checked Sales Orders table if it doesn't already exist
        /// </summary>
        public bool AddSalesOrderCheck(string salesOrder, string excelPath)
        {
            if (string.IsNullOrEmpty(salesOrder) || string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return false;

            // Extract last 5 digits of sales order (matching original VBA logic)
            string soNumber = salesOrder.Length >= 5 ? salesOrder.Substring(salesOrder.Length - 5) : salesOrder;

            // Don't add if already exists
            if (CheckSalesOrder(salesOrder, excelPath))
                return true;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return false;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Checked Sales Orders" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Checked Sales Orders", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                bool success = false;
                if (worksheet != null)
                {
                    // Find next available row
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    int nextRow = lastRow + 1;
                    
                    // Get next ID (column A)
                    int nextId = lastRow; // Since we start from row 2, this gives us the right ID
                    
                    // Add new entry
                    worksheet.Cells[nextRow, 1].Value = nextId; // ID column
                    worksheet.Cells[nextRow, 2].Value = soNumber; // SO column
                    
                    success = true;
                }

                workbook.Save();
                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return success;
            }
            catch
            {
                return false; // If Excel writing fails
            }
        }

        /// <summary>
        /// Get list of all checked sales orders from Excel
        /// </summary>
        public List<string> GetCheckedSalesOrders(string excelPath)
        {
            var checkedOrders = new List<string>();
            
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return checkedOrders;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return checkedOrders;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Checked Sales Orders" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Checked Sales Orders", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                if (worksheet != null)
                {
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var soValue = GetCellValue(worksheet, row, 2)?.ToString()?.Trim(); // Column B = SO
                        if (!string.IsNullOrEmpty(soValue))
                        {
                            checkedOrders.Add(soValue);
                        }
                    }
                }

                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
            catch
            {
                // If Excel reading fails, return empty list
            }

            return checkedOrders;
        }

        /// <summary>
        /// Check if a part has been programmed in the Press Programs table (150 press only)
        /// </summary>
        public bool CheckPartProgrammed(string formDetail, string excelPath)
        {
            if (string.IsNullOrEmpty(formDetail) || string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return false;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return false;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Press Programs" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Press Programs", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                bool isProgrammed = false;
                if (worksheet != null)
                {
                    // Find Form Detail column (should be column A) and 150 Programmed column (should be column B)
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var cellFormDetail = GetCellValue(worksheet, row, 1)?.ToString()?.Trim(); // Column A = Form Detail
                        if (string.Equals(cellFormDetail, formDetail, StringComparison.OrdinalIgnoreCase))
                        {
                            var programmedValue = GetCellValue(worksheet, row, 2); // Column B = 150 Programmed
                            int programmedCount = 0;
                            if (programmedValue != null && int.TryParse(programmedValue.ToString(), out programmedCount))
                            {
                                isProgrammed = programmedCount > 0;
                            }
                            break;
                        }
                    }
                }

                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return isProgrammed;
            }
            catch
            {
                return false; // If Excel reading fails, assume not programmed
            }
        }

        /// <summary>
        /// Add or update a part's programming status in the Press Programs table
        /// </summary>
        public bool AddPartProgram(string formDetail, int programCount, string excelPath)
        {
            if (string.IsNullOrEmpty(formDetail) || string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return false;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return false;

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Press Programs" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Press Programs", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                bool success = false;
                if (worksheet != null)
                {
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    bool found = false;
                    
                    // Look for existing entry
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var cellFormDetail = GetCellValue(worksheet, row, 1)?.ToString()?.Trim(); // Column A = Form Detail
                        if (string.Equals(cellFormDetail, formDetail, StringComparison.OrdinalIgnoreCase))
                        {
                            // Update existing entry
                            worksheet.Cells[row, 2].Value = programCount; // Column B = 150 Programmed
                            found = true;
                            success = true;
                            break;
                        }
                    }
                    
                    // If not found, add new entry
                    if (!found)
                    {
                        int newRow = lastRow + 1;
                        worksheet.Cells[newRow, 1].Value = formDetail; // Form Detail
                        worksheet.Cells[newRow, 2].Value = programCount; // 150 Programmed
                        worksheet.Cells[newRow, 3].Value = 0; // 40 Programmed (legacy, always 0)
                        worksheet.Cells[newRow, 4].Value = 0; // 40 Robot Programmed (legacy, always 0)
                        success = true;
                    }
                }

                workbook.Save();
                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                return success;
            }
            catch
            {
                return false; // If Excel writing fails
            }
        }

        /// <summary>
        /// Check programming status for multiple parts and return missing programs
        /// Optimized version that opens Excel once for all checks
        /// </summary>
        public List<string> GetMissingPrograms(List<string> formDetails, string excelPath)
        {
            var missingPrograms = new List<string>();
            
            if (formDetails == null || formDetails.Count == 0 || string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                return missingPrograms;

            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return formDetails; // If Excel not available, assume all missing

                dynamic excel = Activator.CreateInstance(excelType)!;
                excel.Visible = false;
                excel.DisplayAlerts = false;

                dynamic workbook = excel.Workbooks.Open(excelPath);
                
                // Look for "Press Programs" worksheet
                dynamic? worksheet = null;
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals("Press Programs", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                if (worksheet != null)
                {
                    // Create a dictionary for fast lookup of programmed parts
                    var programmedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    
                    int lastRow = worksheet.UsedRange.Rows.Count;
                    for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                    {
                        var cellFormDetail = GetCellValue(worksheet, row, 1)?.ToString()?.Trim(); // Column A = Form Detail
                        if (!string.IsNullOrEmpty(cellFormDetail))
                        {
                            var programmedValue = GetCellValue(worksheet, row, 2); // Column B = 150 Programmed
                            int programmedCount = 0;
                            if (programmedValue != null && int.TryParse(programmedValue.ToString() ?? "", out programmedCount) && programmedCount > 0)
                            {
                                programmedParts.Add(cellFormDetail);
                            }
                        }
                    }
                    
                    // Check each requested part against the programmed parts
                    foreach (var formDetail in formDetails)
                    {
                        if (!programmedParts.Contains(formDetail))
                        {
                            missingPrograms.Add(formDetail);
                        }
                    }
                }
                else
                {
                    // If worksheet not found, assume all parts are missing programs
                    missingPrograms.AddRange(formDetails);
                }

                workbook.Close(false);
                excel.Quit();

                // Clean up COM objects
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
            catch
            {
                // If Excel reading fails, assume all parts are missing programs
                missingPrograms.AddRange(formDetails);
            }

            return missingPrograms;
        }
        // Helper method for reading Excel cell values
        private object GetCellValue(dynamic worksheet, int row, int col)
        {
            try
            {
                var cell = worksheet.Cells[row, col];
                return cell?.Value;
            }
            catch
            {
                return null;
            }
        }
    }
}