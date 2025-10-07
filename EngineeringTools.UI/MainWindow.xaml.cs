using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.Sqlite;
using System.Data;
using System.ComponentModel;
using XMLIndexer; // Reference the XMLIndexer project

namespace EngineeringTools.UI
{
    public partial class MainWindow : Window
    {
        private readonly string DatabasePath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
        private readonly string ConnectionString;
        private readonly MrpDataManager _mrpManager;
        private string _lastUsedExcelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
        
        // Data collections for binding
        public ObservableCollection<MrpDataItem> MrpData { get; set; }
        public ObservableCollection<CutlistItem> CutlistData { get; set; }
        public ObservableCollection<ComponentItem> ComponentData { get; set; }
        public ObservableCollection<XmlFileItem> XmlFileData { get; set; }
        public ObservableCollection<SalesOrderItem> CheckedSalesOrderData { get; set; }
        public ObservableCollection<MissingProgramItem> MissingProgramData { get; set; }
        
        public MainWindow()
        {
            var debugFile = @"c:\Scripts\EngineeringTools\cutlist_debug.txt";
            File.AppendAllText(debugFile, $"\n=== CONSTRUCTOR START === {DateTime.Now}\n");
            
            InitializeComponent();
            ConnectionString = $"Data Source={DatabasePath}";
            _mrpManager = new MrpDataManager(DatabasePath);
            
            // Import real data from both Excel files
            _mrpManager.ImportRealData();
            
            // Initialize data collections
            MrpData = new ObservableCollection<MrpDataItem>();
            CutlistData = new ObservableCollection<CutlistItem>();
            ComponentData = new ObservableCollection<ComponentItem>();
            XmlFileData = new ObservableCollection<XmlFileItem>();
            CheckedSalesOrderData = new ObservableCollection<SalesOrderItem>();
            MissingProgramData = new ObservableCollection<MissingProgramItem>();
            
            File.AppendAllText(debugFile, "Collections initialized\n");
            
            // Try to find and bind UI elements after InitializeComponent
            try
            {
                var cutlistGrid = this.FindName("CutlistDataGrid") as DataGrid;
                var excelStatus = this.FindName("ExcelStatusTextBlock") as TextBlock;
                var checkedCount = this.FindName("CheckedOrderCountTextBlock") as TextBlock;
                
                File.AppendAllText(debugFile, $"UI Elements found - CutlistGrid: {cutlistGrid != null}, ExcelStatus: {excelStatus != null}, CheckedCount: {checkedCount != null}\n");
                
                if (cutlistGrid != null)
                {
                    cutlistGrid.ItemsSource = CutlistData;
                    File.AppendAllText(debugFile, "CutlistDataGrid bound to CutlistData\n");
                    
                    // Add test item immediately
                    var testItem = new CutlistItem
                    {
                        ComponentPartNumber = "TEST-123",
                        ComponentDescription = "Test Component",
                        TotalQuantity = 1,
                        Material = "Steel",
                        Thickness = "0.125",
                        MaxX = "10.0",
                        MaxY = "5.0",
                        RawMaterialNumber = "RM-001",
                        XmlSource = "Test.xml"
                    };
                    CutlistData.Add(testItem);
                    File.AppendAllText(debugFile, $"Test item added - Collection count: {CutlistData.Count}\n");
                }
                
                // Update Excel status
                if (excelStatus != null && checkedCount != null)
                {
                    string excelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
                    if (File.Exists(excelPath))
                    {
                        excelStatus.Text = "Connected";
                        excelStatus.Foreground = System.Windows.Media.Brushes.Green;
                        checkedCount.Text = "999"; // Test value
                        File.AppendAllText(debugFile, "Excel status updated to Connected\n");
                    }
                    else
                    {
                        excelStatus.Text = "File Not Found";
                        excelStatus.Foreground = System.Windows.Media.Brushes.Red;
                        checkedCount.Text = "0";
                        File.AppendAllText(debugFile, $"Excel file not found at: {excelPath}\n");
                    }
                }
                
                // Update Programming status manually
                File.AppendAllText(debugFile, "Calling UpdateProgramStatus from constructor\n");
                UpdateProgramStatus();
            }
            catch (Exception ex)
            {
                File.AppendAllText(debugFile, $"Constructor binding error: {ex.Message}\n");
            }
            
            File.AppendAllText(debugFile, "=== CONSTRUCTOR END ===\n");
            
            // Skip UI binding for now since UI elements might not exist
        }
        
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            var debugFile = @"c:\Scripts\EngineeringTools\cutlist_debug.txt";
            File.AppendAllText(debugFile, $"\n=== MAIN WINDOW LOADED === {DateTime.Now}\n");
            
            // Simple initialization with automatic Excel import
            try
            {
                // Try to find UI elements manually
                var cutlistGrid = this.FindName("CutlistDataGrid") as DataGrid;
                var mrpGrid = this.FindName("MrpDataGrid") as DataGrid;
                var excelStatus = this.FindName("ExcelStatusTextBlock") as TextBlock;
                
                File.AppendAllText(debugFile, $"CutlistDataGrid found: {cutlistGrid != null}\n");
                File.AppendAllText(debugFile, $"MrpDataGrid found: {mrpGrid != null}\n");
                File.AppendAllText(debugFile, $"ExcelStatusTextBlock found: {excelStatus != null}\n");
                
                // Bind DataGrids to their data sources
                if (mrpGrid != null)
                {
                    mrpGrid.ItemsSource = MrpData;
                    File.AppendAllText(debugFile, "MRP DataGrid bound successfully\n");
                    
                    // Force initial refresh
                    mrpGrid.Items.Refresh();
                }
                else
                {
                    File.AppendAllText(debugFile, "MRP DataGrid not found!\n");
                }
                
                if (cutlistGrid != null)
                {
                    cutlistGrid.ItemsSource = CutlistData;
                    File.AppendAllText(debugFile, "Cutlist DataGrid bound successfully\n");
                    
                    // Add a test item to verify DataGrid works
                    var testItem = new CutlistItem
                    {
                        ComponentPartNumber = "TEST-123",
                        ComponentDescription = "Test Component",
                        TotalQuantity = 1,
                        Material = "Steel",
                        Thickness = "0.125",
                        MaxX = "10.0",
                        MaxY = "5.0", 
                        RawMaterialNumber = "RM-001",
                        XmlSource = "Test.xml"
                    };
                    CutlistData.Add(testItem);
                    File.AppendAllText(debugFile, $"Test item added. CutlistData count: {CutlistData.Count}\n");
                }
                
                // Update Excel status manually
                if (excelStatus != null)
                {
                    string excelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
                    if (File.Exists(excelPath))
                    {
                        excelStatus.Text = "Connected";
                        excelStatus.Foreground = System.Windows.Media.Brushes.Green;
                        File.AppendAllText(debugFile, "Excel status set to Connected\n");
                    }
                    else
                    {
                        excelStatus.Text = "Not Found";
                        excelStatus.Foreground = System.Windows.Media.Brushes.Red;
                        File.AppendAllText(debugFile, "Excel file not found\n");
                    }
                }
                
                // Check database connection silently
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                connection.Close();
                
                // Auto-import Excel data on startup to ensure latest data
                try { StatusTextBlock.Text = "Loading MRP data from Excel..."; } catch { }
                AutoLoadDefaultExcelData();
            }
            catch (Exception ex)
            {
                // Graceful fallback if there are issues
                try { StatusTextBlock.Text = $"Ready (Excel auto-load warning: {ex.Message})"; } catch { }
            }
        }

        /// <summary>
        /// Quick initialization without heavy Excel loading
        /// </summary>
        private void InitializeApplication()
        {
            CheckDatabaseConnection();
            LoadRecentJobs();
            UpdateSalesOrderStatus(); // Initialize Sales Order Check status
            UpdateProgramStatus(); // Initialize Programming Check status
            // AutoLoadDefaultExcelData(); // Data is now loaded during app startup
        }
        
        private void CheckDatabaseConnection()
        {
            try
            {
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                // Check if database has data
                using var command = new SqliteCommand("SELECT COUNT(*) FROM XMLFiles", connection);
                var count = Convert.ToInt32(command.ExecuteScalar());
                
                var dbStatusText = this.FindName("DatabaseStatusTextBlock") as TextBlock;
                var mainStatusText = this.FindName("StatusTextBlock") as TextBlock;
                
                if (dbStatusText != null)
                {
                    dbStatusText.Text = $"Database: Connected ({count:N0} XML files)";
                    dbStatusText.Foreground = System.Windows.Media.Brushes.Green;
                }
                if (mainStatusText != null)
                {
                    mainStatusText.Text = "Database connection established.";
                }
            }
            catch (Exception ex)
            {
                var dbStatusText = this.FindName("DatabaseStatusTextBlock") as TextBlock;
                var mainStatusText = this.FindName("StatusTextBlock") as TextBlock;
                
                if (dbStatusText != null)
                {
                    dbStatusText.Text = "Database: Connection Failed";
                    dbStatusText.Foreground = System.Windows.Media.Brushes.Red;
                }
                if (mainStatusText != null)
                {
                    mainStatusText.Text = $"Database error: {ex.Message}";
                }
                MessageBox.Show($"Database connection failed: {ex.Message}", "Database Error", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        
        private void LoadRecentJobs()
        {
            // TODO: Load recent jobs from JobPartNumber table
            StatusTextBlock.Text = "Ready - Enter job number to begin";
        }
        
        // Menu and Toolbar Event Handlers
        private void NewJob_Click(object sender, RoutedEventArgs e)
        {
            JobNumberTextBox.Clear();
            PartNumberTextBox.Clear();
            JobDescriptionTextBox.Clear();
            MrpData.Clear();
            CutlistData.Clear();
            StatusTextBlock.Text = "New job started - Enter job details";
        }
        
        private void OpenJob_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new JobSearchDialog();
            if (dialog.ShowDialog() == true)
            {
                LoadJobData(dialog.SelectedJobNumber);
            }
        }
        
        private void LoadJob_Click(object sender, RoutedEventArgs e)
        {
            var jobNumber = JobNumberTextBox.Text.Trim();
            if (string.IsNullOrEmpty(jobNumber))
            {
                MessageBox.Show("Please enter a job number.", "Input Required", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            LoadJobData(jobNumber);
        }
        
        private void PreviewCutlist_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var jobNumber = JobNumberTextBox.Text.Trim();
                if (string.IsNullOrEmpty(jobNumber))
                {
                    MessageBox.Show("Please enter a job number to preview its cutlist.", "Input Required", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                // Check if it's an I-job or a make job (H, J, K, etc. - everything except I)
                bool isIJob = jobNumber.StartsWith("I", StringComparison.OrdinalIgnoreCase);
                bool isMakeJob = !isIJob; // All non-I jobs are make jobs (H, J, K, L, etc.)
                
                // For both I-jobs and standard jobs, try to load the job data and preview
                LoadJobData(jobNumber);
                
                if (MrpData.Count > 0)
                {
                    // Generate cutlist automatically
                    GenerateCutlist();
                    
                    // Ensure MRP data grid is properly bound and refreshed
                    var mrpGrid = this.FindName("MrpDataGrid") as DataGrid;
                    if (mrpGrid != null)
                    {
                        mrpGrid.ItemsSource = null;
                        mrpGrid.ItemsSource = MrpData;
                        mrpGrid.Items.Refresh();
                    }
                    
                    string jobTypeMessage = isIJob ? "I-Job (Established Cutlist)" : "Make Job (H/J/K+ series)";
                    MessageBox.Show($"Cutlist preview loaded for {jobTypeMessage}: {jobNumber}\n\n" +
                                   $"Found {MrpData.Count} MRP items, Generated {CutlistData.Count} cutlist components\n" +
                                   "You can now view the cutlist in the Generated Cutlist section below.",
                                   "Cutlist Preview", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"No data found for job {jobNumber}\n\n" +
                                   "This could mean:\n" +
                                   "• Job number doesn't exist in the MRP data\n" +
                                   "• Job has not been loaded into the system\n" +
                                   "• Check the job number spelling",
                                   "No Job Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error previewing cutlist: {ex.Message}", "Preview Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void LoadFromExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Show dialog to select Excel file
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Select MRP Excel File",
                    Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*",
                    DefaultExt = ".xls",
                    FileName = "Priority List Master SHOP-SQL.xls"
                };
                
                if (dialog.ShowDialog() == true)
                {
                    LoadMrpFromExcel(dialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel file: {ex.Message}", "Excel Load Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void LoadMrpFromExcel(string excelPath)
        {
            try
            {
                // StatusTextBlock.Text = "Loading MRP data from Excel...";
                
                var success = _mrpManager.ImportFromExcel(excelPath);
                
                if (success)
                {
                    // Store the Excel file path for later use in SaveJob
                    _lastUsedExcelPath = excelPath;
                    
                    MessageBox.Show($"MRP data imported successfully from Excel!\n\nFile: {Path.GetFileName(excelPath)}", 
                        "Import Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    // StatusTextBlock.Text = "MRP data imported from Excel - Ready to load jobs";
                }
                else
                {
                    MessageBox.Show("No data was imported from the Excel file.", "Import Warning", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to import MRP data: {ex.Message}", "Import Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                
                // StatusTextBlock.Text = "Excel import failed";
            }
        }
        
        /// <summary>
        /// Automatically load MRP data from the default Excel file on startup
        /// </summary>
        private void AutoLoadDefaultExcelData()
        {
            try
            {
                // Try to load from the default Excel file location
                string defaultExcelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
                
                Console.WriteLine($"Auto-loading Excel data from: {defaultExcelPath}");
                
                if (File.Exists(defaultExcelPath))
                {
                    Console.WriteLine("Excel file found, importing data...");
                    var success = _mrpManager.ImportFromExcel(defaultExcelPath);
                    
                    if (success)
                    {
                        // Store the Excel file path for later use in SaveJob
                        _lastUsedExcelPath = defaultExcelPath;
                        
                        Console.WriteLine("Excel data auto-loaded successfully!");
                        // Silent success - no popup, just update status if available
                        try { StatusTextBlock.Text = "Ready - MRP data auto-loaded from Excel"; } catch { }
                    }
                    else
                    {
                        Console.WriteLine("Excel import failed");
                        try { StatusTextBlock.Text = "Ready - Excel auto-load failed, use Load from Excel button"; } catch { }
                    }
                }
                else
                {
                    Console.WriteLine("Excel file not found at default location");
                    try { StatusTextBlock.Text = "Ready - Use Load from Excel button to import data"; } catch { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Auto-load Excel error: {ex.Message}");
                try { StatusTextBlock.Text = "Ready - Excel auto-load error, use Load from Excel button"; } catch { }
            }
        }
        
        private void LoadJobData(string jobNumber)
        {
            try
            {
                StatusTextBlock.Text = $"Loading job {jobNumber}...";
                
                // TODO: This is where we would integrate with MRP system
                // For now, simulate MRP data lookup
                LoadSimulatedMrpData(jobNumber);
                
                StatusTextBlock.Text = $"Job {jobNumber} loaded - {MrpData.Count} parts found";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Error loading job data";
                MessageBox.Show($"Error loading job: {ex.Message}", "Load Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void LoadSimulatedMrpData(string jobNumber)
        {
            // Load real MRP data from database
            MrpData.Clear();

            try
            {
                var mrpItems = _mrpManager.GetMrpDataForJob(jobNumber);

                // DEBUG: Show what we found
                string debugMsg = $"DEBUG: Job Search Results\n\n";
                debugMsg += $"Searching for job: '{jobNumber}'\n";
                debugMsg += $"Excel file: {Path.GetFileName(_lastUsedExcelPath)}\n";
                debugMsg += $"MRP items found: {mrpItems.Count}\n\n";
                
                if (mrpItems.Count > 0)
                {
                    debugMsg += "Found parts:\n";
                    foreach (var item in mrpItems.Take(5))
                    {
                        debugMsg += $"- Job: '{item.JobNumber}', Part: '{item.PartNumber}', Rev: '{item.Revision}', Desc: '{item.Description}'\n";
                    }
                    if (mrpItems.Count > 5) debugMsg += $"... and {mrpItems.Count - 5} more";
                }
                else
                {
                    debugMsg += "No parts found for this job.";
                }
                
                MessageBox.Show(debugMsg, "Job Search Debug", MessageBoxButton.OK, MessageBoxImage.Information);

                if (mrpItems.Count == 0)
                {
                    // Job not found - provide clear feedback
                    MessageBox.Show($"Job '{jobNumber}' was not found in the MRP data.\n\nPlease verify the job number is correct (case-insensitive search used).\n\nData source: {Path.GetFileName(_lastUsedExcelPath)}", 
                        "Job Not Found", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                foreach (var item in mrpItems)
                {
                    var mrpDataItem = new MrpDataItem
                    {
                        Job = item.JobNumber, // Use the actual job number from database, not the search term
                        PartNumber = item.PartNumber,
                        Revision = item.Revision,
                        Quantity = item.Quantity,
                        Description = item.Description,
                        XmlStatus = item.XmlStatus,
                        HighestRelease = item.HighestRelease > 0 ? $"REL{item.HighestRelease}" : "N/A"
                    };
                    
                    // Check XML availability for this part
                    CheckXmlAvailability(mrpDataItem);
                    
                    MrpData.Add(mrpDataItem);
                }
            }
            catch (Exception ex)
            {
                // No fallback data - show clear error message
                MessageBox.Show($"Error loading MRP data: {ex.Message}\n\nPlease check:\n• Excel file path: {_lastUsedExcelPath}\n• File is not open in Excel\n• File permissions", 
                    "MRP Data Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw; // Re-throw to ensure caller knows about the error
            }
        }        private void CheckXmlAvailability(MrpDataItem item)
        {
            try
            {
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                // Find XML files for this part
                using var command = new SqliteCommand(@"
                    SELECT PartNumber, Revision, Release, COUNT(*) as ComponentCount
                    FROM XMLFiles xf
                    LEFT JOIN Components c ON xf.ID = c.XMLFileID
                    WHERE xf.PartNumber LIKE @partNumber
                    GROUP BY xf.PartNumber, xf.Revision, xf.Release
                    ORDER BY CAST(xf.Release AS INTEGER) DESC
                    LIMIT 1", connection);
                
                command.Parameters.AddWithValue("@partNumber", $"%{item.PartNumber}%");
                
                using var reader = command.ExecuteReader();
                if (reader.Read())
                {
                    item.XmlStatus = "Available";
                    item.HighestRelease = $"REL{reader["Release"]}";
                }
                else
                {
                    // Try a broader search to see what's in the database
                    reader.Close();
                    using var debugCmd = new SqliteCommand("SELECT PartNumber FROM XMLFiles WHERE PartNumber LIKE @debugSearch LIMIT 5", connection);
                    debugCmd.Parameters.AddWithValue("@debugSearch", $"%{item.PartNumber.Substring(0, Math.Min(4, item.PartNumber.Length))}%");
                    using var debugReader = debugCmd.ExecuteReader();
                    
                    var similarParts = new List<string>();
                    while (debugReader.Read())
                    {
                        similarParts.Add(debugReader["PartNumber"]?.ToString() ?? "");
                    }
                    
                    // DEBUG: Show popup for first item to debug
                    if (item == MrpData.FirstOrDefault())
                    {
                        string debugMsg = $"DEBUG INFO for job search:\n\n";
                        debugMsg += $"Searched for part: '{item.PartNumber}'\n";
                        debugMsg += $"Job: '{item.Job}'\n";
                        debugMsg += $"Description: '{item.Description}'\n";
                        debugMsg += $"Revision: '{item.Revision}'\n\n";
                        debugMsg += $"XML Search results:\n";
                        if (similarParts.Count > 0)
                        {
                            debugMsg += $"Similar parts found:\n{string.Join("\n", similarParts)}";
                        }
                        else
                        {
                            debugMsg += "No similar parts found in XML database";
                        }
                        
                        MessageBox.Show(debugMsg, "Debug Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    
                    item.XmlStatus = "Not Found";
                    item.HighestRelease = "N/A";
                }
            }
            catch
            {
                item.XmlStatus = "Error";
                item.HighestRelease = "N/A";
            }
        }
        
        private void GenerateCutlist_Click(object sender, RoutedEventArgs e)
        {
            if (MrpData.Count == 0)
            {
                MessageBox.Show("No MRP data loaded. Please load a job first.", "No Data", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            GenerateCutlist();
        }
        
        private void GenerateCutlist()
        {
            try
            {
                StatusTextBlock.Text = "Generating cutlist...";
                CutlistData.Clear();
                
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                // Get the job number to determine job type
                var jobNumber = MrpData.FirstOrDefault()?.Job ?? "";
                bool isIJob = jobNumber.StartsWith("I", StringComparison.OrdinalIgnoreCase);
                
                if (isIJob)
                {
                    // For I-jobs: Use highest release of the specific revision from established XMLs
                    foreach (var mrpItem in MrpData)
                    {
                        LoadCutlistForIJob(connection, mrpItem);
                    }
                }
                else
                {
                    // For make jobs (H/J/K+ series): Only process items with available XML
                    foreach (var mrpItem in MrpData.Where(m => m.XmlStatus == "Available"))
                    {
                        LoadCutlistForPart(connection, mrpItem);
                    }
                }
                
                StatusTextBlock.Text = $"Cutlist generated - {CutlistData.Count} components from {MrpData.Count} MRP items";
                
                // Update database status using FindName
                try
                {
                    using var dbConnection = new SqliteConnection(ConnectionString);
                    dbConnection.Open();
                    var dbStatusText = this.FindName("DatabaseStatusTextBlock") as TextBlock;
                    if (dbStatusText != null)
                    {
                        dbStatusText.Text = "Database: Connected";
                        dbStatusText.Foreground = System.Windows.Media.Brushes.Green;
                    }
                    dbConnection.Close();
                }
                catch
                {
                    var dbStatusText = this.FindName("DatabaseStatusTextBlock") as TextBlock;
                    if (dbStatusText != null)
                    {
                        dbStatusText.Text = "Database: Error";
                        dbStatusText.Foreground = System.Windows.Media.Brushes.Red;
                    }
                }
                
                // Switch to Job Management tab to show results
                MainTabControl.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Error generating cutlist";
                MessageBox.Show($"Error generating cutlist: {ex.Message}", "Generation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void LoadCutlistForPart(SqliteConnection connection, MrpDataItem mrpItem)
        {
            // Improved cutlist generation with better deduplication and null filtering
            using var command = new SqliteCommand(@"
                WITH HighestRelease AS (
                    SELECT MAX(CAST(xf.Release AS INTEGER)) as MaxRelease
                    FROM Components c
                    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
                    WHERE c.ParentPartNumber LIKE @partNumber 
                        AND c.ComponentType = 'Make'
                ),
                BestComponents AS (
                    SELECT 
                        c.ComponentPartNumber,
                        c.ComponentDescription,
                        c.TotalQuantity,
                        c.Material,
                        c.Thickness,
                        c.MaxX,
                        c.MaxY,
                        c.RawMaterialNumber,
                        xf.FileName as XmlSource,
                        ROW_NUMBER() OVER (
                            PARTITION BY c.ComponentPartNumber 
                            ORDER BY 
                                CASE WHEN c.MaxX IS NOT NULL AND c.MaxY IS NOT NULL THEN 1 ELSE 2 END,
                                CASE WHEN c.RawMaterialNumber IS NOT NULL AND c.RawMaterialNumber != '' THEN 1 ELSE 2 END,
                                c.TotalQuantity DESC,
                                CAST(xf.Release AS INTEGER) DESC
                        ) as rn
                    FROM Components c
                    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
                    JOIN HighestRelease hr ON CAST(xf.Release AS INTEGER) = hr.MaxRelease
                    WHERE c.ParentPartNumber LIKE @partNumber 
                        AND c.ComponentType = 'Make'
                        AND c.ComponentPartNumber IS NOT NULL
                        AND c.ComponentPartNumber != ''
                        AND c.ComponentPartNumber != 'null'
                        AND c.ComponentPartNumber NOT LIKE '%.asm'
                        AND c.ComponentPartNumber NOT LIKE '%.sldasm'
                        AND c.ComponentPartNumber NOT LIKE '%.ASM'
                        AND c.ComponentPartNumber NOT LIKE '%.SLDASM'
                )
                SELECT 
                    ComponentPartNumber,
                    ComponentDescription,
                    TotalQuantity,
                    COALESCE(Material, '') as Material,
                    COALESCE(Thickness, '') as Thickness,
                    COALESCE(MaxX, '') as MaxX,
                    COALESCE(MaxY, '') as MaxY,
                    COALESCE(RawMaterialNumber, '') as RawMaterialNumber,
                    XmlSource
                FROM BestComponents
                WHERE rn = 1
                ORDER BY ComponentPartNumber", connection);
            
            command.Parameters.AddWithValue("@partNumber", $"%{mrpItem.PartNumber}%");
            
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                // Strip file extensions from component part number
                var rawPartNumber = reader["ComponentPartNumber"]?.ToString() ?? "";
                var cleanPartNumber = System.IO.Path.GetFileNameWithoutExtension(rawPartNumber);
                if (string.IsNullOrEmpty(cleanPartNumber))
                    cleanPartNumber = rawPartNumber; // Fallback if no extension
                
                var cutlistItem = new CutlistItem
                {
                    ComponentPartNumber = cleanPartNumber,
                    ComponentDescription = reader["ComponentDescription"]?.ToString() ?? "",
                    TotalQuantity = Convert.ToInt32(reader["TotalQuantity"]) * mrpItem.Quantity,
                    Material = reader["Material"]?.ToString() ?? "",
                    Thickness = reader["Thickness"]?.ToString() ?? "",
                    MaxX = reader["MaxX"]?.ToString() ?? "",
                    MaxY = reader["MaxY"]?.ToString() ?? "",
                    RawMaterialNumber = reader["RawMaterialNumber"]?.ToString() ?? "",
                    XmlSource = reader["XmlSource"]?.ToString() ?? ""
                };
                
                CutlistData.Add(cutlistItem);
            }
        }
        
        private void LoadCutlistForIJob(SqliteConnection connection, MrpDataItem mrpItem)
        {
            // For I-jobs: Find the highest release of the specific revision for established cutlists
            
            // DEBUG: Show what parameters we're searching with
            var partNumberParam = $"%{mrpItem.PartNumber}%";
            var revisionParam = mrpItem.Revision != null ? mrpItem.Revision.Trim() + "%" : "%";
            
            var debugFile = @"c:\Scripts\EngineeringTools\cutlist_debug.txt";
            File.AppendAllText(debugFile, $"\n=== CUTLIST DEBUG FOR I-JOB === {DateTime.Now}\n");
            File.AppendAllText(debugFile, $"Job: {mrpItem.Job}\n");
            File.AppendAllText(debugFile, $"Part: {mrpItem.PartNumber}\n");
            File.AppendAllText(debugFile, $"Revision: '{mrpItem.Revision}'\n");
            File.AppendAllText(debugFile, $"Search PartNumber: '{partNumberParam}'\n");
            File.AppendAllText(debugFile, $"Search Revision: '{revisionParam}'\n");
            File.AppendAllText(debugFile, "================================\n");
            
            using var command = new SqliteCommand(@"
                WITH HighestReleaseOfRevision AS (
                    SELECT MAX(CAST(xf.Release AS INTEGER)) as MaxRelease
                    FROM XMLFiles xf
                    WHERE xf.PartNumber LIKE @partNumber 
                        AND (xf.Revision LIKE @revisionLike)
                ),
                BestComponents AS (
                    SELECT 
                        c.ComponentPartNumber,
                        c.ComponentDescription,
                        c.TotalQuantity,
                        c.Material,
                        c.Thickness,
                        c.MaxX,
                        c.MaxY,
                        c.RawMaterialNumber,
                        xf.FileName as XmlSource,
                        xf.Revision,
                        xf.Release,
                        ROW_NUMBER() OVER (
                            PARTITION BY c.ComponentPartNumber 
                            ORDER BY 
                                CASE WHEN c.MaxX IS NOT NULL AND c.MaxY IS NOT NULL THEN 1 ELSE 2 END,
                                CASE WHEN c.RawMaterialNumber IS NOT NULL AND c.RawMaterialNumber != '' THEN 1 ELSE 2 END,
                                c.TotalQuantity DESC,
                                CAST(xf.Release AS INTEGER) DESC
                        ) as rn
                    FROM Components c
                    JOIN XMLFiles xf ON c.XMLFileID = xf.ID
                    JOIN HighestReleaseOfRevision hr ON CAST(xf.Release AS INTEGER) = hr.MaxRelease
                    WHERE xf.PartNumber LIKE @partNumber 
                        AND (xf.Revision LIKE @revisionLike)
                        AND c.ComponentType = 'Make'
                        AND c.ComponentPartNumber IS NOT NULL
                        AND c.ComponentPartNumber != ''
                        AND c.ComponentPartNumber != 'null'
                        AND c.ComponentPartNumber NOT LIKE '%.asm'
                        AND c.ComponentPartNumber NOT LIKE '%.sldasm'
                        AND c.ComponentPartNumber NOT LIKE '%.ASM'
                        AND c.ComponentPartNumber NOT LIKE '%.SLDASM'
                )
                SELECT 
                    ComponentPartNumber,
                    ComponentDescription,
                    TotalQuantity,
                    COALESCE(Material, '') as Material,
                    COALESCE(Thickness, '') as Thickness,
                    COALESCE(MaxX, '') as MaxX,
                    COALESCE(MaxY, '') as MaxY,
                    COALESCE(RawMaterialNumber, '') as RawMaterialNumber,
                    XmlSource,
                    Revision,
                    Release
                FROM BestComponents
                WHERE rn = 1
                ORDER BY ComponentPartNumber", connection);
            
            command.Parameters.AddWithValue("@partNumber", partNumberParam);
            command.Parameters.AddWithValue("@revisionLike", revisionParam);
            
            File.AppendAllText(debugFile, $"Executing query with parameters:\n");
            File.AppendAllText(debugFile, $"  @partNumber = '{partNumberParam}'\n");
            File.AppendAllText(debugFile, $"  @revisionLike = '{revisionParam}'\n");
            
            using var reader = command.ExecuteReader();
            int componentCount = 0;
            while (reader.Read())
            {
                componentCount++;
                
                // Strip file extensions from component part number
                var rawPartNumber = reader["ComponentPartNumber"]?.ToString() ?? "";
                var cleanPartNumber = System.IO.Path.GetFileNameWithoutExtension(rawPartNumber);
                if (string.IsNullOrEmpty(cleanPartNumber))
                    cleanPartNumber = rawPartNumber; // Fallback if no extension
                
                var cutlistItem = new CutlistItem
                {
                    ComponentPartNumber = cleanPartNumber,
                    ComponentDescription = reader["ComponentDescription"]?.ToString() ?? "",
                    TotalQuantity = Convert.ToInt32(reader["TotalQuantity"]) * mrpItem.Quantity,
                    Material = reader["Material"]?.ToString() ?? "",
                    Thickness = reader["Thickness"]?.ToString() ?? "",
                    MaxX = reader["MaxX"]?.ToString() ?? "",
                    MaxY = reader["MaxY"]?.ToString() ?? "",
                    RawMaterialNumber = reader["RawMaterialNumber"]?.ToString() ?? "",
                    XmlSource = $"{reader["XmlSource"]} (Rev: {reader["Revision"]}, Rel: {reader["Release"]})"
                };
                
                File.AppendAllText(debugFile, $"Adding component: {cutlistItem.ComponentPartNumber} - {cutlistItem.ComponentDescription}\n");
                CutlistData.Add(cutlistItem);
            }
            
            File.AppendAllText(debugFile, $"Found {componentCount} components for I-job cutlist\n");
            File.AppendAllText(debugFile, $"Total CutlistData count: {CutlistData.Count}\n");
            File.AppendAllText(debugFile, "=== END CUTLIST DEBUG ===\n");
            
            // Force collection refresh
            try
            {
                System.Windows.Data.CollectionViewSource.GetDefaultView(CutlistData).Refresh();
            }
            catch (Exception refreshEx)
            {
                File.AppendAllText(debugFile, $"Refresh error: {refreshEx.Message}\n");
            }
        }
        
        private void CreateBurnList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Check if we have an Excel file path
                if (string.IsNullOrEmpty(_lastUsedExcelPath) || !File.Exists(_lastUsedExcelPath))
                {
                    MessageBox.Show("No Excel file loaded. Please load MRP data from Excel first.", 
                        "No Excel File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // Get jobs ready for burning (engineering completed, XML files assigned in Excel)
                var jobsReadyForBurning = _mrpManager.GetJobsReadyForBurning(_lastUsedExcelPath);
                
                if (jobsReadyForBurning.Count == 0)
                {
                    MessageBox.Show("No jobs are ready for burning. Jobs must have XML files assigned in Excel.", 
                        "No Jobs Ready", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                // Show dialog to select burn list type
                var result = MessageBox.Show(
                    $"Found {jobsReadyForBurning.Count} jobs ready for burning.\n\n" +
                    "Click 'Yes' for ERP format (.erp)\n" +
                    "Click 'No' for WOL format (.wol)\n" +
                    "Click 'Cancel' to abort",
                    "Select Burn List Format", 
                    MessageBoxButton.YesNoCancel, 
                    MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Cancel)
                    return;
                
                bool isErpFormat = result == MessageBoxResult.Yes;
                string fileExtension = isErpFormat ? ".erp" : ".wol";
                string fileDescription = isErpFormat ? "ERP Exchange" : "WOL Cutting";
                
                // Generate burn list file
                string outputPath = GenerateBurnListFile(jobsReadyForBurning, isErpFormat);
                
                MessageBox.Show(
                    $"{fileDescription} burn list created successfully!\n\n" +
                    $"File: {outputPath}\n" +
                    $"Jobs processed: {jobsReadyForBurning.Count}",
                    "Burn List Created", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Information);
                    
                // Optionally open the file location
                var openResult = MessageBox.Show(
                    "Would you like to open the file location?",
                    "Open File Location", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);
                    
                if (openResult == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{outputPath}\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating burn list: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private string GenerateBurnListFile(List<XMLIndexer.MrpDataManager.MrpItem> jobs, bool isErpFormat)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string baseDir = @"C:\Scripts\EngineeringTools\xmlCutlist";
            
            // Ensure directory exists
            Directory.CreateDirectory(baseDir);
            
            string fileName = isErpFormat ? "cutlist.erp" : "BurnListSigma.wol";
            string backupFileName = isErpFormat ? $"erp_backup_{timestamp}.erp" : $"sig_backup_{timestamp}.txt";
            
            string outputPath = Path.Combine(baseDir, fileName);
            string backupPath = Path.Combine(baseDir, backupFileName);
            
            if (isErpFormat)
            {
                GenerateErpFile(jobs, outputPath, backupPath);
            }
            else
            {
                GenerateWolFile(jobs, outputPath, backupPath);
            }
            
            return outputPath;
        }
        
        private void GenerateErpFile(List<XMLIndexer.MrpDataManager.MrpItem> jobs, string outputPath, string backupPath)
        {
            using var writer = new StreamWriter(outputPath);
            using var backup = new StreamWriter(backupPath);
            
            // Write XML headers
            string xmlHeader = "<ErpExchange>";
            string ordersOpen = "\t<Orders>";
            
            writer.WriteLine(xmlHeader);
            backup.WriteLine(xmlHeader);
            writer.WriteLine(ordersOpen);
            backup.WriteLine(ordersOpen);
            
            foreach (var job in jobs)
            {
                // Check if engineering is really completed
                if (!_mrpManager.IsEngineeringCompleted(job.JobNumber, _lastUsedExcelPath))
                    continue;
                
                // Get XML file path
                string xmlFilePath = _mrpManager.GetXmlFilePathForJob(job.JobNumber, _lastUsedExcelPath);
                if (string.IsNullOrEmpty(xmlFilePath))
                    continue;
                
                // Generate ERP order XML
                string orderXml = GenerateErpOrderXml(job, xmlFilePath);
                
                writer.WriteLine(orderXml);
                backup.WriteLine(orderXml);
            }
            
            // Write XML footers
            string ordersClose = "\t</Orders>";
            string xmlFooter = "</ErpExchange>";
            
            writer.WriteLine(ordersClose);
            backup.WriteLine(ordersClose);
            writer.WriteLine(xmlFooter);
            backup.WriteLine(xmlFooter);
        }
        
        private void GenerateWolFile(List<XMLIndexer.MrpDataManager.MrpItem> jobs, string outputPath, string backupPath)
        {
            using var writer = new StreamWriter(outputPath);
            using var backup = new StreamWriter(backupPath);
            
            // Write WOL headers (format based on Sigma cutting equipment)
            writer.WriteLine("# WOL Burn List Generated: " + DateTime.Now);
            backup.WriteLine("# WOL Burn List Generated: " + DateTime.Now);
            
            foreach (var job in jobs)
            {
                // Check if engineering is really completed
                if (!_mrpManager.IsEngineeringCompleted(job.JobNumber, _lastUsedExcelPath))
                    continue;
                
                // Get XML file path
                string xmlFilePath = _mrpManager.GetXmlFilePathForJob(job.JobNumber, _lastUsedExcelPath);
                if (string.IsNullOrEmpty(xmlFilePath))
                    continue;
                
                // Generate WOL entry
                string wolEntry = GenerateWolEntry(job, xmlFilePath);
                
                writer.WriteLine(wolEntry);
                backup.WriteLine(wolEntry);
            }
        }
        
        private string GenerateErpOrderXml(XMLIndexer.MrpDataManager.MrpItem job, string xmlFilePath)
        {
            var xml = new System.Text.StringBuilder();
            
            xml.AppendLine("\t\t<ErpOrder>");
            xml.AppendLine("\t\t\t<ImportType>NewOrder</ImportType>");
            xml.AppendLine($"\t\t\t<OrderNumber>{job.JobNumber.Trim().ToUpper()}</OrderNumber>");
            
            // Parse dates - use today if not available
            string startDate = DateTime.Now.ToString("yyyy-MM-dd");
            string targetDate = DateTime.Now.AddDays(7).ToString("yyyy-MM-dd");
            
            if (!string.IsNullOrEmpty(job.DueDate))
            {
                if (DateTime.TryParse(job.DueDate, out DateTime dueDate))
                {
                    targetDate = dueDate.ToString("yyyy-MM-dd");
                }
            }
            
            xml.AppendLine($"\t\t\t<StartDate>{startDate}</StartDate>");
            xml.AppendLine($"\t\t\t<TargetDate>{targetDate}</TargetDate>");
            xml.AppendLine("\t\t\t<ProductionStrategy>MaterialAdministrationOrder</ProductionStrategy>");
            xml.AppendLine("\t\t\t<Automatic>True</Automatic>");
            
            // Add job details as XML sub-elements
            xml.AppendLine($"\t\t\t<PartNumber>{job.PartNumber}</PartNumber>");
            xml.AppendLine($"\t\t\t<Quantity>{job.Quantity}</Quantity>");
            xml.AppendLine($"\t\t\t<XmlFile>{xmlFilePath}</XmlFile>");
            
            xml.AppendLine("\t\t</ErpOrder>");
            
            return xml.ToString();
        }
        
        private string GenerateWolEntry(XMLIndexer.MrpDataManager.MrpItem job, string xmlFilePath)
        {
            // WOL format for Sigma cutting equipment
            return $"{job.JobNumber}\t{job.PartNumber}\t{job.Quantity}\t{xmlFilePath}\t{job.Description}";
        }
        
        private void SaveJob_Click(object sender, RoutedEventArgs e)
        {
            if (MrpData.Count == 0)
            {
                MessageBox.Show("No MRP data loaded. Please load a job first.", "No Data", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            if (CutlistData.Count == 0)
            {
                MessageBox.Show("No cutlist generated. Please generate a cutlist first.", "No Cutlist", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Sales Order Validation (matching original VBA workflow)
            string jobNumber = MrpData.FirstOrDefault()?.Job ?? "";
            bool isIJob = jobNumber.StartsWith("I", StringComparison.OrdinalIgnoreCase);
            
            if (!isIJob)
            {
                // For make jobs (H/J/K+ series), require sales order validation
                var salesOrderResult = MessageBox.Show(
                    $"Sales Order Validation Required\n\n" +
                    $"Job Number: {jobNumber}\n\n" +
                    $"Before saving this job, you must verify that the associated sales order has been checked.\n\n" +
                    $"• Go to the 'Sales Order Check' tab to verify the sales order\n" +
                    $"• Enter the sales order number and mark it as checked\n" +
                    $"• Then return here to save the job\n\n" +
                    $"Has the sales order for this job been verified and checked?",
                    "Sales Order Validation Required",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (salesOrderResult == MessageBoxResult.No)
                {
                    MessageBox.Show(
                        "Job save cancelled. Please verify the sales order in the 'Sales Order Check' tab first.\n\n" +
                        "This verification step ensures quality control and matches the original engineering workflow.",
                        "Sales Order Verification Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }
            }
            else
            {
                // I-jobs skip sales order validation (established cutlists don't need validation)
                MessageBox.Show(
                    $"I-Job Detected: {jobNumber}\n\n" +
                    "Sales order validation is not required for I-jobs (established cutlists).\nProceeding with job save...",
                    "I-Job - Sales Order Check Skipped",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            // Programming Validation (for all jobs - check if make items are programmed for 150 press)
            try
            {
                string excelPath = @"C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx";
                var makeItems = CutlistData.Select(c => c.ComponentPartNumber).Distinct().ToList();
                var missingPrograms = _mrpManager.GetMissingPrograms(makeItems, excelPath);
                
                if (missingPrograms.Count > 0)
                {
                    var programResult = MessageBox.Show(
                        $"Programming Validation Warning\n\n" +
                        $"Job Number: {jobNumber}\n\n" +
                        $"The following {missingPrograms.Count} part(s) are missing 150 press programs:\n\n" +
                        $"{string.Join("\n", missingPrograms.Take(10))}" +
                        $"{(missingPrograms.Count > 10 ? $"\n... and {missingPrograms.Count - 10} more" : "")}\n\n" +
                        $"• These parts should be programmed before completing the job\n" +
                        $"• Use the 'Programming Check' tab to review and update programming status\n" +
                        $"• Note: Only 150 press programming is checked (40 press is discontinued)\n\n" +
                        $"Do you want to continue saving despite missing programs?",
                        "Programming Validation Warning",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (programResult == MessageBoxResult.No)
                    {
                        MessageBox.Show(
                            "Job save cancelled. Please ensure all make items are programmed for 150 press.\n\n" +
                            "Use the 'Programming Check' tab to:\n" +
                            "• Check individual part programming status\n" +
                            "• Mark parts as programmed when completed\n" +
                            "• View missing programs for the current job\n\n" +
                            "This verification step ensures manufacturing readiness.",
                            "Programming Verification Required",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show(
                        $"Programming Check: ✓ All parts programmed\n\n" +
                        $"All {makeItems.Count} make items in this job are programmed for 150 press.\n" +
                        $"Proceeding with job save...",
                        "Programming Check Passed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                var continueResult = MessageBox.Show(
                    $"Programming validation error: {ex.Message}\n\n" +
                    "Unable to verify programming status. Do you want to continue saving the job anyway?",
                    "Programming Check Error",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                    
                if (continueResult == MessageBoxResult.No)
                    return;
            }
            
            // Show print dialog options
            var result = MessageBox.Show(
                "Do you want to print the cutlist?\n\nYes = Print and save record\nNo = Just save record\nCancel = Cancel", 
                "Print Cutlist", 
                MessageBoxButton.YesNoCancel, 
                MessageBoxImage.Question);
            
            if (result == MessageBoxResult.Cancel)
                return;
                
            try
            {
                if (result == MessageBoxResult.Yes)
                {
                    PrintCutlist();
                }
                
                UpdateExcelWithJobAssignments();
                
                string message = result == MessageBoxResult.Yes ? 
                    "Cutlist printed and Excel updated!" : 
                    "Excel updated with job assignments!";
                    
                MessageBox.Show(message, "Success", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void UpdateExcelWithJobAssignments()
        {
            try
            {
                // Use the stored Excel file path from the last load operation
                if (string.IsNullOrEmpty(_lastUsedExcelPath) || !File.Exists(_lastUsedExcelPath))
                {
                    throw new Exception($"Excel file not found: {_lastUsedExcelPath ?? "No file specified"}");
                }
                
                // Prepare job/XML assignments dictionary
                var jobXmlAssignments = new Dictionary<string, string>();
                
                foreach (var mrpItem in MrpData.Where(m => m.XmlStatus == "Available"))
                {
                    // Find the primary XML file for this part (most common one in cutlist)
                    var primaryXmlFile = CutlistData
                        .Where(c => !string.IsNullOrEmpty(c.XmlSource))
                        .GroupBy(c => c.XmlSource)
                        .OrderByDescending(g => g.Count())
                        .FirstOrDefault()?.Key;
                    
                    if (!string.IsNullOrEmpty(primaryXmlFile) && !string.IsNullOrEmpty(mrpItem.Job))
                    {
                        jobXmlAssignments[mrpItem.Job] = primaryXmlFile;
                    }
                }
                
                if (jobXmlAssignments.Count > 0)
                {
                    // Update Excel file with job/XML assignments
                    bool success = _mrpManager.UpdateExcelWithJobXmlAssignments(_lastUsedExcelPath, jobXmlAssignments);
                    
                    if (!success)
                    {
                        throw new Exception("No job assignments were updated in Excel");
                    }
                }
            }
            catch (Exception ex)
            {
                // Don't throw - just log the error, Excel update is secondary to database save
                MessageBox.Show($"Warning: Could not update Excel file: {ex.Message}", 
                    "Excel Update Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        
        private void PrintCutlist()
        {
            try
            {
                // Create a formatted cutlist report
                var cutlistReport = GenerateCutlistReport();
                
                // Save to temp file and open for printing
                var tempFile = Path.Combine(Path.GetTempPath(), $"Cutlist_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllText(tempFile, cutlistReport);
                
                // Open with default text editor for printing
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = tempFile,
                    UseShellExecute = true
                });
                
                // StatusTextBlock.Text = $"Cutlist report generated: {Path.GetFileName(tempFile)}";
            }
            catch (Exception ex)
            {
                throw new Exception($"Print failed: {ex.Message}");
            }
        }
        
        private string GenerateCutlistReport()
        {
            var report = new System.Text.StringBuilder();
            var mrpItem = MrpData.FirstOrDefault();
            
            // Header
            report.AppendLine("=====================================");
            report.AppendLine("           CUTLIST REPORT");
            report.AppendLine("=====================================");
            report.AppendLine($"Job Number: {mrpItem?.Job ?? "N/A"}");
            report.AppendLine($"Part Number: {mrpItem?.PartNumber ?? "N/A"}");
            report.AppendLine($"Description: {mrpItem?.Description ?? "N/A"}");
            report.AppendLine($"Quantity: {mrpItem?.Quantity ?? 0}");
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine("=====================================");
            report.AppendLine();
            
            // Column headers
            report.AppendLine($"{"Component Part Number",-30} {"Qty",-5} {"Material",-15} {"Thickness",-10} {"MaxX",-8} {"MaxY",-8} {"Raw Material",-15}");
            report.AppendLine(new string('-', 110));
            
            // Components
            foreach (var item in CutlistData.OrderBy(c => c.ComponentPartNumber))
            {
                report.AppendLine($"{item.ComponentPartNumber,-30} {item.TotalQuantity,-5} {item.Material,-15} {item.Thickness,-10} {item.MaxX,-8} {item.MaxY,-8} {item.RawMaterialNumber,-15}");
            }
            
            report.AppendLine();
            report.AppendLine($"Total Components: {CutlistData.Count}");
            
            // XML Sources
            report.AppendLine();
            report.AppendLine("XML Sources:");
            var xmlSources = CutlistData.Where(c => !string.IsNullOrEmpty(c.XmlSource))
                                      .Select(c => c.XmlSource)
                                      .Distinct()
                                      .OrderBy(x => x);
            foreach (var source in xmlSources)
            {
                report.AppendLine($"  - {source}");
            }
            
            return report.ToString();
        }
        
        private void BrowseDatabase_Click(object sender, RoutedEventArgs e)
        {
            MainTabControl.SelectedIndex = 1; // Switch to Database Browser tab
            LoadDatabaseOverview();
        }
        
        private void LoadDatabaseOverview()
        {
            try
            {
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                using var command = new SqliteCommand(@"
                    SELECT 
                        'XMLFiles' as TableName,
                        COUNT(*) as RecordCount,
                        'XML file definitions' as Description
                    FROM XMLFiles
                    UNION ALL
                    SELECT 
                        'Components' as TableName,
                        COUNT(*) as RecordCount,
                        'Component definitions from XMLs' as Description
                    FROM Components
                    UNION ALL
                    SELECT 
                        'PartData' as TableName,
                        COUNT(*) as RecordCount,
                        'Part properties and dimensions' as Description
                    FROM PartData", connection);
                
                var dataTable = new DataTable();
                using var reader = command.ExecuteReader();
                dataTable.Load(reader);
                
                DatabaseResultsDataGrid.ItemsSource = dataTable.DefaultView;
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Error loading database overview";
                MessageBox.Show($"Error: {ex.Message}", "Database Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void DatabaseSearch_Click(object sender, RoutedEventArgs e)
        {
            var searchTerm = DatabaseSearchTextBox.Text.Trim();
            if (string.IsNullOrEmpty(searchTerm))
            {
                LoadDatabaseOverview();
                return;
            }
            
            SearchDatabase(searchTerm);
        }
        
        private void SearchDatabase(string searchTerm)
        {
            try
            {
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                using var command = new SqliteCommand(@"
                    SELECT 
                        xf.PartNumber,
                        xf.Revision,
                        xf.Release,
                        xf.FileName,
                        COUNT(c.ID) as ComponentCount,
                        xf.FileModifiedDate
                    FROM XMLFiles xf
                    LEFT JOIN Components c ON xf.ID = c.XMLFileID
                    WHERE xf.PartNumber LIKE @search 
                        OR xf.FileName LIKE @search
                    GROUP BY xf.ID
                    ORDER BY xf.PartNumber, CAST(xf.Release AS INTEGER) DESC", connection);
                
                command.Parameters.AddWithValue("@search", $"%{searchTerm}%");
                
                var dataTable = new DataTable();
                using var reader = command.ExecuteReader();
                dataTable.Load(reader);
                
                DatabaseResultsDataGrid.ItemsSource = dataTable.DefaultView;
                StatusTextBlock.Text = $"Found {dataTable.Rows.Count} matching records";
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Search error";
                MessageBox.Show($"Search error: {ex.Message}", "Search Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void DatabaseClear_Click(object sender, RoutedEventArgs e)
        {
            DatabaseSearchTextBox.Clear();
            LoadDatabaseOverview();
        }
        
        private void AnalyzeXml_Click(object sender, RoutedEventArgs e)
        {
            var partTextBox = this.FindName("XmlAnalysisPartTextBox") as TextBox;
            if (partTextBox == null)
            {
                MessageBox.Show("UI element not found. Please restart the application.", "UI Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            var partNumber = partTextBox.Text.Trim();
            if (string.IsNullOrEmpty(partNumber))
            {
                MessageBox.Show("Please enter a part number to analyze.", "Input Required", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            AnalyzeXmlForPart(partNumber);
        }
        
        private void AnalyzeXmlForPart(string partNumber)
        {
            try
            {
                ComponentData.Clear();
                XmlFileData.Clear();
                
                using var connection = new SqliteConnection(ConnectionString);
                connection.Open();
                
                // Load components
                LoadComponentsForPart(connection, partNumber);
                
                // Load XML files
                LoadXmlFilesForPart(connection, partNumber);
                
                // Update status using FindName
                var statusText = this.FindName("XmlAnalysisStatus") as TextBlock;
                if (statusText != null)
                {
                    statusText.Text = $"Found {ComponentData.Count} components, {XmlFileData.Count} XML files";
                }
                
                // Bind data to grids using FindName
                var componentsGrid = this.FindName("XmlComponentsDataGrid") as DataGrid;
                var xmlFilesGrid = this.FindName("XmlFilesDataGrid") as DataGrid;
                
                if (componentsGrid != null)
                {
                    componentsGrid.ItemsSource = null;
                    componentsGrid.ItemsSource = ComponentData;
                    componentsGrid.Items.Refresh();
                }
                
                if (xmlFilesGrid != null)
                {
                    xmlFilesGrid.ItemsSource = null;
                    xmlFilesGrid.ItemsSource = XmlFileData;
                    xmlFilesGrid.Items.Refresh();
                }
                
                // Switch to XML Analysis tab using FindName
                var mainTabControl = this.FindName("MainTabControl") as TabControl;
                if (mainTabControl != null)
                {
                    mainTabControl.SelectedIndex = 2; // XML Analysis tab
                }
            }
            catch (Exception ex)
            {
                var statusText = this.FindName("XmlAnalysisStatus") as TextBlock;
                if (statusText != null)
                {
                    statusText.Text = "Analysis failed";
                }
                MessageBox.Show($"Analysis error: {ex.Message}", "Analysis Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void LoadComponentsForPart(SqliteConnection connection, string partNumber)
        {
            using var command = new SqliteCommand(@"
                SELECT 
                    c.ComponentPartNumber,
                    c.ComponentType,
                    c.TotalQuantity,
                    c.Material,
                    c.Thickness,
                    c.AssemblyLevel
                FROM Components c
                JOIN XMLFiles xf ON c.XMLFileID = xf.ID
                WHERE c.ParentPartNumber LIKE @partNumber
                ORDER BY c.ComponentPartNumber", connection);
            
            command.Parameters.AddWithValue("@partNumber", $"%{partNumber}%");
            
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                ComponentData.Add(new ComponentItem
                {
                    ComponentPartNumber = reader["ComponentPartNumber"]?.ToString() ?? "",
                    ComponentType = reader["ComponentType"]?.ToString() ?? "",
                    TotalQuantity = Convert.ToInt32(reader["TotalQuantity"]),
                    Material = reader["Material"]?.ToString() ?? "",
                    Thickness = reader["Thickness"]?.ToString() ?? "",
                    AssemblyLevel = Convert.ToInt32(reader["AssemblyLevel"])
                });
            }
        }
        
        private void LoadXmlFilesForPart(SqliteConnection connection, string partNumber)
        {
            using var command = new SqliteCommand(@"
                SELECT 
                    xf.PartNumber,
                    xf.Revision,
                    xf.Release,
                    xf.FileName,
                    xf.FileModifiedDate
                FROM XMLFiles xf
                WHERE xf.PartNumber LIKE @partNumber
                ORDER BY xf.PartNumber, CAST(xf.Release AS INTEGER) DESC", connection);
            
            command.Parameters.AddWithValue("@partNumber", $"%{partNumber}%");
            
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                XmlFileData.Add(new XmlFileItem
                {
                    PartNumber = reader["PartNumber"]?.ToString() ?? "",
                    Revision = reader["Revision"]?.ToString() ?? "",
                    Release = reader["Release"]?.ToString() ?? "",
                    FileName = reader["FileName"]?.ToString() ?? "",
                    FileModifiedDate = reader["FileModifiedDate"]?.ToString() ?? ""
                });
            }
        }
        
        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            CheckDatabaseConnection();
            var statusTextBlock = this.FindName("StatusTextBlock") as TextBlock;
            if (statusTextBlock != null)
            {
                statusTextBlock.Text = "Interface refreshed";
            }
        }
        
        private void RebuildIndex_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("This will rebuild the XML index database. This may take several minutes. Continue?", 
                "Rebuild Index", MessageBoxButton.YesNo, MessageBoxImage.Question);
            
            if (result == MessageBoxResult.Yes)
            {
                // TODO: Call XMLIndexer rebuild process
                MessageBox.Show("Index rebuild functionality will be implemented next.", "Coming Soon", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        private void XmlSearch_Click(object sender, RoutedEventArgs e)
        {
            var mainTabControl = this.FindName("MainTabControl") as TabControl;
            if (mainTabControl != null)
            {
                mainTabControl.SelectedIndex = 2; // Switch to XML Analysis tab
            }
        }
        
        private void CutlistGenerator_Click(object sender, RoutedEventArgs e)
        {
            var mainTabControl = this.FindName("MainTabControl") as TabControl;
            if (mainTabControl != null)
            {
                mainTabControl.SelectedIndex = 0; // Switch to Job Management tab
            }
        }
        
        private void JobHistory_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement job history dialog
            MessageBox.Show("Job history functionality will be implemented next.", "Coming Soon", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        private void About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Engineering Tools .NET\nModern Engineering Workflow Management\n\nVersion 1.0\nBuilt on SQLite Database", 
                "About Engineering Tools", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Sales Order Check Event Handlers
        private void CheckSalesOrder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var salesOrderTextBox = this.FindName("SalesOrderTextBox") as TextBox;
                var statusTextBlock = this.FindName("SalesOrderStatusTextBlock") as TextBlock;
                
                if (salesOrderTextBox == null || statusTextBlock == null)
                {
                    MessageBox.Show("UI elements not found. Please restart the application.", "UI Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                var salesOrder = salesOrderTextBox.Text.Trim();
                if (string.IsNullOrEmpty(salesOrder))
                {
                    statusTextBlock.Text = "Please enter a sales order number";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // Show loading message
                statusTextBlock.Text = "Checking sales order status...";
                statusTextBlock.Foreground = System.Windows.Media.Brushes.Blue;
                
                // Force UI update
                this.UpdateLayout();

                // Use fast database lookup instead of slow Excel COM
                bool isChecked = _mrpManager.CheckSalesOrderInDatabase(salesOrder);

                if (isChecked)
                {
                    statusTextBlock.Text = $"✓ Sales Order {salesOrder} has been checked";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Green;
                }
                else
                {
                    statusTextBlock.Text = $"⚠ Sales Order {salesOrder} has NOT been checked";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                }

                UpdateSalesOrderStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking sales order: {ex.Message}", "Sales Order Check Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddSalesOrder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var salesOrderTextBox = this.FindName("SalesOrderTextBox") as TextBox;
                var statusTextBlock = this.FindName("SalesOrderStatusTextBlock") as TextBlock;
                
                if (salesOrderTextBox == null || statusTextBlock == null)
                {
                    MessageBox.Show("UI elements not found. Please restart the application.", "UI Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                var salesOrder = salesOrderTextBox.Text.Trim();
                if (string.IsNullOrEmpty(salesOrder))
                {
                    statusTextBlock.Text = "Please enter a sales order number";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // Use fast database operations instead of slow Excel COM
                
                // Check if it already exists
                if (_mrpManager.CheckSalesOrderInDatabase(salesOrder))
                {
                    statusTextBlock.Text = $"Sales Order {salesOrder} is already checked";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Blue;
                    return;
                }

                // Add to checked list
                bool success = _mrpManager.AddSalesOrderToDatabase(salesOrder);
                
                if (success)
                {
                    statusTextBlock.Text = $"✓ Sales Order {salesOrder} has been marked as checked";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Green;
                    
                    // Clear the textbox and refresh the list
                    salesOrderTextBox.Clear();
                    RefreshCheckedOrders_Click(sender, e);
                }
                else
                {
                    statusTextBlock.Text = $"Failed to mark Sales Order {salesOrder} as checked";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                }

                UpdateSalesOrderStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding sales order: {ex.Message}", "Sales Order Add Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshCheckedOrders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string excelPath = @"C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx";
                var checkedOrders = _mrpManager.GetCheckedSalesOrders(excelPath);
                
                CheckedSalesOrderData.Clear();
                
                // Take last 50 orders and reverse to show most recent first
                var recentOrders = checkedOrders.TakeLast(50).Reverse();
                
                foreach (var order in recentOrders)
                {
                    CheckedSalesOrderData.Add(new SalesOrderItem
                    {
                        SalesOrder = order,
                        DateChecked = "Recently", // Would need to store actual dates
                        Status = "Checked"
                    });
                }

                UpdateSalesOrderStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing checked orders: {ex.Message}", "Refresh Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportCheckedOrders_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Export to CSV functionality will be implemented in a future update.", 
                           "Feature Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ImportCheckedOrders_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Import from CSV functionality will be implemented in a future update.", 
                           "Feature Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ClearAllOrders_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to clear ALL checked sales orders? This cannot be undone.", 
                                        "Clear All Orders", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            
            if (result == MessageBoxResult.Yes)
            {
                MessageBox.Show("Clear all functionality will be implemented in a future update.", 
                               "Feature Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void UpdateSalesOrderStatus()
        {
            var debugFile = @"c:\Scripts\EngineeringTools\cutlist_debug.txt";
            File.AppendAllText(debugFile, $"\n=== SALES ORDER STATUS UPDATE === {DateTime.Now}\n");
            
            try
            {
                var excelStatusTextBlock = this.FindName("ExcelStatusTextBlock") as TextBlock;
                var checkedOrderCountTextBlock = this.FindName("CheckedOrderCountTextBlock") as TextBlock;
                
                string excelPath = @"C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx";
                File.AppendAllText(debugFile, $"Checking Excel file: {excelPath}\n");
                File.AppendAllText(debugFile, $"File exists: {File.Exists(excelPath)}\n");
                
                if (File.Exists(excelPath))
                {
                    try 
                    { 
                        if (excelStatusTextBlock != null)
                        {
                            excelStatusTextBlock.Text = "Connected";
                            excelStatusTextBlock.Foreground = System.Windows.Media.Brushes.Green;
                            File.AppendAllText(debugFile, "Successfully set status to Connected\n");
                        }
                    }
                    catch (Exception uiEx)
                    {
                        File.AppendAllText(debugFile, $"UI update error: {uiEx.Message}\n");
                    }
                    
                    try
                    {
                        var checkedOrders = _mrpManager.GetCheckedSalesOrdersFromDatabase();
                        if (checkedOrderCountTextBlock != null)
                        {
                            checkedOrderCountTextBlock.Text = checkedOrders.Count.ToString();
                            File.AppendAllText(debugFile, $"Checked orders count: {checkedOrders.Count}\n");
                        }
                    }
                    catch (Exception countEx)
                    {
                        File.AppendAllText(debugFile, $"Count update error: {countEx.Message}\n");
                    }
                }
                else
                {
                    try
                    {
                        if (excelStatusTextBlock != null)
                        {
                            excelStatusTextBlock.Text = "Not Found";
                            excelStatusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                        }
                        if (checkedOrderCountTextBlock != null)
                        {
                            checkedOrderCountTextBlock.Text = "0";
                        }
                        File.AppendAllText(debugFile, "Excel file not found\n");
                    }
                    catch (Exception uiEx)
                    {
                        File.AppendAllText(debugFile, $"UI error for 'not found': {uiEx.Message}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    var excelStatusTextBlock = this.FindName("ExcelStatusTextBlock") as TextBlock;
                    var checkedOrderCountTextBlock = this.FindName("CheckedOrderCountTextBlock") as TextBlock;
                    
                    if (excelStatusTextBlock != null)
                    {
                        excelStatusTextBlock.Text = "Error";
                        excelStatusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                    }
                    if (checkedOrderCountTextBlock != null)
                    {
                        checkedOrderCountTextBlock.Text = "0";
                    }
                }
                catch { }
                File.AppendAllText(debugFile, $"Error in UpdateSalesOrderStatus: {ex.Message}\n");
            }
            
            File.AppendAllText(debugFile, "=== END SALES ORDER STATUS UPDATE ===\n");
        }

        // Programming Check Event Handlers
        private void CheckProgram_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var programPartTextBox = this.FindName("ProgramPartTextBox") as TextBox;
                var statusTextBlock = this.FindName("ProgramStatusTextBlock") as TextBlock;
                
                if (programPartTextBox == null || statusTextBlock == null)
                {
                    MessageBox.Show("UI elements not found. Please restart the application.", "UI Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                var partNumber = programPartTextBox.Text.Trim();
                if (string.IsNullOrEmpty(partNumber))
                {
                    statusTextBlock.Text = "Please enter a part number";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // Show loading message
                statusTextBlock.Text = "Checking programming status...";
                statusTextBlock.Foreground = System.Windows.Media.Brushes.Blue;
                
                // Force UI update
                this.UpdateLayout();

                // Use fast database lookup instead of slow Excel COM
                bool isProgrammed = _mrpManager.CheckPartProgrammedInDatabase(partNumber);

                if (isProgrammed)
                {
                    statusTextBlock.Text = $"✓ Part {partNumber} is programmed for 150 press";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Green;
                }
                else
                {
                    statusTextBlock.Text = $"⚠ Part {partNumber} is NOT programmed for 150 press";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                }

                UpdateProgramStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking programming: {ex.Message}", "Programming Check Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddProgram_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var programPartTextBox = this.FindName("ProgramPartTextBox") as TextBox;
                var statusTextBlock = this.FindName("ProgramStatusTextBlock") as TextBlock;
                
                if (programPartTextBox == null || statusTextBlock == null)
                {
                    MessageBox.Show("UI elements not found. Please restart the application.", "UI Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                var partNumber = programPartTextBox.Text.Trim();
                if (string.IsNullOrEmpty(partNumber))
                {
                    statusTextBlock.Text = "Please enter a part number";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // Use fast database operations instead of slow Excel COM
                
                // Check if it's already programmed
                if (_mrpManager.CheckPartProgrammedInDatabase(partNumber))
                {
                    statusTextBlock.Text = $"Part {partNumber} is already programmed";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Blue;
                    return;
                }

                // Add to programmed list
                _mrpManager.AddProgrammedPartToDatabase(partNumber);
                
                statusTextBlock.Text = $"✓ Part {partNumber} has been marked as programmed";
                statusTextBlock.Foreground = System.Windows.Media.Brushes.Green;
                
                // Clear the textbox and refresh status
                programPartTextBox.Clear();
                UpdateProgramStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding program: {ex.Message}", "Programming Add Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CheckJobPrograms_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var programJobTextBox = this.FindName("ProgramJobTextBox") as TextBox;
                var statusTextBlock = this.FindName("ProgramStatusTextBlock") as TextBlock;
                
                if (programJobTextBox == null || statusTextBlock == null)
                {
                    MessageBox.Show("UI elements not found. Please restart the application.", "UI Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                var jobNumber = programJobTextBox.Text.Trim();
                if (string.IsNullOrEmpty(jobNumber))
                {
                    statusTextBlock.Text = "Please enter a job number to check programming";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Red;
                    return;
                }

                // Show loading message
                statusTextBlock.Text = "Loading job parts and checking programming...";
                statusTextBlock.Foreground = System.Windows.Media.Brushes.Blue;
                
                // Force UI update
                this.UpdateLayout();

                // Get parts for this job from the database
                var jobParts = GetJobPartNumbers(jobNumber);
                
                if (jobParts.Count == 0)
                {
                    statusTextBlock.Text = $"No parts found for job {jobNumber}. Load the job first or check job number.";
                    statusTextBlock.Foreground = System.Windows.Media.Brushes.Orange;
                    return;
                }
                
                // Use fast database method instead of slow Excel COM
                var missingPrograms = _mrpManager.GetMissingProgramsFromDatabase(jobParts);
                
                MissingProgramData.Clear();
                
                // Add missing programs to the display
                foreach (var partNumber in missingPrograms)
                {
                    MissingProgramData.Add(new MissingProgramItem
                    {
                        PartNumber = partNumber,
                        DateFound = DateTime.Now.ToString("yyyy-MM-dd"),
                        JobNumber = jobNumber,
                        Status = "Missing Program"
                    });
                }
                
                // Force refresh the DataGrid
                var missingProgramsGrid = this.FindName("MissingProgramsDataGrid") as DataGrid;
                if (missingProgramsGrid != null)
                {
                    missingProgramsGrid.ItemsSource = null;
                    missingProgramsGrid.ItemsSource = MissingProgramData;
                    missingProgramsGrid.Items.Refresh();
                }

                statusTextBlock.Text = $"Job {jobNumber}: Found {MissingProgramData.Count} parts missing programs (from {jobParts.Count} total parts)";
                statusTextBlock.Foreground = MissingProgramData.Count > 0 ? 
                    System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.Green;

                UpdateProgramStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking job programs: {ex.Message}", "Job Program Check Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Get all part numbers for a specific job from cutlist data or MRP data
        /// </summary>
        private List<string> GetJobPartNumbers(string jobNumber)
        {
            var partNumbers = new List<string>();
            
            // First try from current cutlist if it matches the job
            if (CutlistData.Count > 0)
            {
                var currentJobFromCutlist = MrpData.FirstOrDefault()?.Job ?? "";
                if (currentJobFromCutlist.Equals(jobNumber, StringComparison.OrdinalIgnoreCase))
                {
                    partNumbers.AddRange(CutlistData.Select(c => c.ComponentPartNumber).Distinct());
                    return partNumbers;
                }
            }
            
            // If not in current cutlist, try from MRP data
            var jobMrpItems = MrpData.Where(m => m.Job.Equals(jobNumber, StringComparison.OrdinalIgnoreCase));
            if (jobMrpItems.Any())
            {
                partNumbers.AddRange(jobMrpItems.Select(m => m.PartNumber).Distinct());
                return partNumbers;
            }
            
            // If still not found, try to generate cutlist for this job
            try
            {
                LoadJobData(jobNumber);
                if (MrpData.Count > 0)
                {
                    GenerateCutlist();
                    partNumbers.AddRange(CutlistData.Select(c => c.ComponentPartNumber).Distinct());
                }
            }
            catch
            {
                // Silent failure - will return empty list
            }
            
            return partNumbers.Distinct().ToList();
        }

        private void RefreshPrograms_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string excelPath = @"C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx";
                
                // For refresh, get all programmed parts and find what's missing from Press Programs
                var allProgrammedParts = _mrpManager.GetProgrammedParts(excelPath);
                // Get a sample list of parts to check - this would ideally come from a database or file
                var allKnownParts = new List<string>(); // TODO: Replace with actual source of all part numbers
                var missingPrograms = _mrpManager.GetMissingPrograms(allKnownParts, excelPath);
                
                MissingProgramData.Clear();
                
                // Take last 50 missing programs
                var recentMissing = missingPrograms.TakeLast(50).Reverse();
                
                foreach (var partNumber in recentMissing)
                {
                    MissingProgramData.Add(new MissingProgramItem
                    {
                        PartNumber = partNumber,
                        DateFound = "Recently",
                        JobNumber = "Various",
                        Status = "Missing Program"
                    });
                }

                UpdateProgramStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing programs: {ex.Message}", "Refresh Error", 
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportMissingPrograms_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Export missing programs functionality will be implemented in a future update.", 
                           "Feature Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ImportProgrammedParts_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Import programmed parts functionality will be implemented in a future update.", 
                           "Feature Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GenerateProgramReport_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Generate program report functionality will be implemented in a future update.", 
                           "Feature Coming Soon", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void UpdateProgramStatus()
        {
            var debugFile = @"c:\Scripts\EngineeringTools\cutlist_debug.txt";
            File.AppendAllText(debugFile, $"\n=== PROGRAM STATUS UPDATE === {DateTime.Now}\n");
            
            try
            {
                // Use FindName to locate UI elements
                var programExcelStatus = this.FindName("ProgramExcelStatusTextBlock") as TextBlock;
                var programmedPartCount = this.FindName("ProgrammedPartCountTextBlock") as TextBlock;
                
                File.AppendAllText(debugFile, $"ProgramExcelStatusTextBlock found: {programExcelStatus != null}\n");
                File.AppendAllText(debugFile, $"ProgrammedPartCountTextBlock found: {programmedPartCount != null}\n");
                
                string excelPath = @"C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx";
                File.AppendAllText(debugFile, $"Checking Excel file for programming: {excelPath}\n");
                File.AppendAllText(debugFile, $"File exists: {File.Exists(excelPath)}\n");
                
                if (File.Exists(excelPath))
                {
                    File.AppendAllText(debugFile, "Excel file found, getting programmed parts...\n");
                    
                    try
                    {
                        var programmedParts = _mrpManager.GetProgrammedPartsFromDatabase();
                        File.AppendAllText(debugFile, $"GetProgrammedParts succeeded, count: {programmedParts.Count}\n");
                        
                        if (programExcelStatus != null && programmedPartCount != null)
                        {
                            programExcelStatus.Text = "Connected";
                            programExcelStatus.Foreground = System.Windows.Media.Brushes.Green;
                            programmedPartCount.Text = programmedParts.Count.ToString();
                            File.AppendAllText(debugFile, "Program status set to Connected\n");
                        }
                        else
                        {
                            File.AppendAllText(debugFile, "Could not find UI elements to update\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(debugFile, $"GetProgrammedParts failed: {ex.Message}\n");
                        if (programExcelStatus != null && programmedPartCount != null)
                        {
                            programExcelStatus.Text = "Error";
                            programExcelStatus.Foreground = System.Windows.Media.Brushes.Red;
                            programmedPartCount.Text = "0";
                        }
                    }
                }
                else
                {
                    if (programExcelStatus != null && programmedPartCount != null)
                    {
                        programExcelStatus.Text = "Not Found";
                        programExcelStatus.Foreground = System.Windows.Media.Brushes.Red;
                        programmedPartCount.Text = "0";
                    }
                    File.AppendAllText(debugFile, "Excel file not found for programming\n");
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(debugFile, $"UpdateProgramStatus error: {ex.Message}\n");
            }
            
            File.AppendAllText(debugFile, "=== END PROGRAM STATUS UPDATE ===\n");
        }
    }
    
    // Data Model Classes
    public class MrpDataItem : INotifyPropertyChanged
    {
        private string _xmlStatus = "";
        private string _highestRelease = "";
        
        public string Job { get; set; } = "";
        public string PartNumber { get; set; } = "";
        public string Revision { get; set; } = "";
        public int Quantity { get; set; }
        public string Description { get; set; } = "";
        
        public string XmlStatus 
        { 
            get => _xmlStatus; 
            set { _xmlStatus = value; OnPropertyChanged(nameof(XmlStatus)); }
        }
        
        public string HighestRelease 
        { 
            get => _highestRelease; 
            set { _highestRelease = value; OnPropertyChanged(nameof(HighestRelease)); }
        }
        
        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
    
    public class CutlistItem
    {
        public string ComponentPartNumber { get; set; } = "";
        public string ComponentDescription { get; set; } = "";
        public int TotalQuantity { get; set; }
        public string Material { get; set; } = "";
        public string Thickness { get; set; } = "";
        public string MaxX { get; set; } = "";
        public string MaxY { get; set; } = "";
        public string RawMaterialNumber { get; set; } = "";
        public string XmlSource { get; set; } = "";
    }
    
    public class ComponentItem
    {
        public string ComponentPartNumber { get; set; } = "";
        public string ComponentType { get; set; } = "";
        public int TotalQuantity { get; set; }
        public string Material { get; set; } = "";
        public string Thickness { get; set; } = "";
        public int AssemblyLevel { get; set; }
    }
    
    public class XmlFileItem
    {
        public string PartNumber { get; set; } = "";
        public string Revision { get; set; } = "";
        public string Release { get; set; } = "";
        public string FileName { get; set; } = "";
        public string FileModifiedDate { get; set; } = "";
    }
    
    // Data model for Sales Order items
    public class SalesOrderItem
    {
        public string SalesOrder { get; set; } = "";
        public string DateChecked { get; set; } = "";
        public string Status { get; set; } = "";
    }
    
    // Data model for Missing Program items
    public class MissingProgramItem
    {
        public string PartNumber { get; set; } = "";
        public string DateFound { get; set; } = "";
        public string JobNumber { get; set; } = "";
        public string Status { get; set; } = "";
    }
    
    // Placeholder for Job Search Dialog
    public class JobSearchDialog : Window
    {
        public string SelectedJobNumber { get; private set; } = "";
        
        public JobSearchDialog()
        {
            Title = "Search Jobs";
            Width = 400;
            Height = 300;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            
            // Simple placeholder implementation
            var button = new Button { Content = "Cancel", Margin = new Thickness(10) };
            button.Click += (s, e) => DialogResult = false;
            Content = button;
        }
    }
}
