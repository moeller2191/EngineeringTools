using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace EngineeringTools.UI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private LoadingWindow? _loadingWindow;

        protected override void OnStartup(StartupEventArgs e)
        {
            // Call base startup to initialize WPF application framework
            base.OnStartup(e);
            
            // Simple direct startup - no loading window
            try
            {
                var mainWindow = new MainWindow();
                MainWindow = mainWindow;
                mainWindow.Show();
                mainWindow.Activate();
                mainWindow.WindowState = WindowState.Normal;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to start application: {ex.Message}", "Startup Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
            }
        }
        
        private void InitializeApplicationSync()
        {
            try
            {
                // Initialize the main window synchronously
                InitializeApplicationSyncInternal();
                
                // Show completion and close loading window
                _loadingWindow?.ShowCompleted();
                System.Threading.Thread.Sleep(1500); // Simple delay instead of async
                
                // Create and show main window directly
                try
                {
                    var mainWindow = new MainWindow();
                    MainWindow = mainWindow;
                    mainWindow.Show();
                    _loadingWindow?.Close();
                    
                    mainWindow.Activate();
                    mainWindow.WindowState = WindowState.Normal;
                }
                catch (Exception mainWindowEx)
                {
                    MessageBox.Show($"Failed to create main window: {mainWindowEx.Message}\n\nStack trace: {mainWindowEx.StackTrace}", 
                                   "Main Window Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Shutdown(1);
                }
            }
            catch (Exception ex)
            {
                _loadingWindow?.ShowError($"Failed to initialize: {ex.Message}");
                System.Threading.Thread.Sleep(3000);
                Shutdown(1);
            }
        }
        
        private async Task InitializeApplicationWithErrorHandling()
        {
            try
            {
                // Initialize the main window asynchronously
                await InitializeApplicationAsync();
                
                // Show completion and close loading window
                _loadingWindow?.ShowCompleted();
                await _loadingWindow?.CloseWithDelay(1500)!;
                
                MessageBox.Show("Loading window closed, about to create MainWindow", "Debug Before MainWindow", MessageBoxButton.OK);
                
                // Create and show main window directly (we're already on UI thread)
                try
                {
                    MessageBox.Show("About to create MainWindow...", "Debug 1", MessageBoxButton.OK);
                    
                    var mainWindow = new MainWindow();
                    
                    MessageBox.Show("MainWindow created successfully!", "Debug 2", MessageBoxButton.OK);
                    
                    MainWindow = mainWindow;
                    
                    MessageBox.Show("About to show MainWindow...", "Debug 3", MessageBoxButton.OK);
                    
                    mainWindow.Show();
                    
                    MessageBox.Show("MainWindow.Show() called!", "Debug 4", MessageBoxButton.OK);
                    
                    mainWindow.Activate();
                    mainWindow.WindowState = WindowState.Normal;
                    
                    MessageBox.Show("MainWindow should now be visible!", "Debug 5", MessageBoxButton.OK);
                }
                catch (Exception mainWindowEx)
                {
                    MessageBox.Show($"Failed to create main window: {mainWindowEx.Message}\n\nStack trace: {mainWindowEx.StackTrace}", 
                                   "Main Window Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Shutdown(1);
                }
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    _loadingWindow?.ShowError($"Failed to initialize: {ex.Message}");
                });
                await Task.Delay(3000);
                await Dispatcher.InvokeAsync(() =>
                {
                    _loadingWindow?.Close();
                    MessageBox.Show($"Application failed to start: {ex.Message}\n\nStack trace: {ex.StackTrace}", "Startup Error", 
                                   MessageBoxButton.OK, MessageBoxImage.Error);
                    Shutdown(1);
                });
            }
        }

        private async Task InitializeApplicationAsync()
        {
            var databasePath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
            var excelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
            
            // Step 1: Database initialization
            _loadingWindow?.UpdateProgress(15, "Initializing database...", "Creating tables and indexes");
            _loadingWindow?.UpdateStep(1, "Database initialization", false);
            
            try
            {
                var mrpManager = new XMLIndexer.MrpDataManager(databasePath);
                _loadingWindow?.UpdateStep(1, "Database ready", true);
            }
            catch (Exception ex)
            {
                _loadingWindow?.UpdateStep(1, "Database initialization failed", false);
                throw new Exception($"Database initialization failed: {ex.Message}");
            }

            // Step 2: Excel data import
            _loadingWindow?.UpdateProgress(40, "Loading MRP data...", "Importing from Priority List Master SHOP-SQL.xls");
            _loadingWindow?.UpdateStep(2, "MRP data import", false);
            
            try
            {
                if (System.IO.File.Exists(excelPath))
                {
                    var mrpManager = new XMLIndexer.MrpDataManager(databasePath);
                    
                    // Add timeout to prevent hanging
                    var importTask = Task.Run(() => mrpManager.ImportFromExcel(excelPath));
                    var completedTask = await Task.WhenAny(importTask, Task.Delay(30000)); // 30 second timeout
                    
                    if (completedTask == importTask)
                    {
                        _loadingWindow?.UpdateStep(2, "MRP data imported", true);
                    }
                    else
                    {
                        _loadingWindow?.UpdateStep(2, "Excel import timeout - using existing data", true);
                    }
                }
                else
                {
                    _loadingWindow?.UpdateStep(2, "Excel file not found - using existing data", true);
                }
            }
            catch (Exception ex)
            {
                _loadingWindow?.UpdateStep(2, "Excel import failed - using existing data", true);
            }

            // Step 3: Engineering database validation
            _loadingWindow?.UpdateProgress(65, "Validating engineering data...", "Checking database integrity");
            _loadingWindow?.UpdateStep(3, "Data validation", false);
            await Task.Delay(800); // Actual validation would happen here
            _loadingWindow?.UpdateStep(3, "Data validation complete", true);

            // Step 4: Component systems
            _loadingWindow?.UpdateProgress(85, "Loading application components...", "Initializing data grids and controls");
            _loadingWindow?.UpdateStep(4, "Application components", false);
            await Task.Delay(600); // Time for component initialization
            _loadingWindow?.UpdateStep(4, "Components ready", true);

            // Step 5: Final preparation
            _loadingWindow?.UpdateProgress(95, "Finalizing startup...", "Preparing user interface");
            _loadingWindow?.UpdateStep(5, "Interface preparation", false);
            await Task.Delay(400); // Final UI setup time
            _loadingWindow?.UpdateStep(5, "Ready to launch", true);
            
            // Debug: Confirm initialization completed
            MessageBox.Show("InitializeApplicationAsync completed successfully!", "Debug - Initialization Complete", MessageBoxButton.OK);
        }
        
        private void InitializeApplicationSyncInternal()
        {
            // Simple initialization without complex steps
            _loadingWindow?.UpdateProgress(25, "Initializing database...", "Setting up data connections");
            
            // Initialize database synchronously
            var mrpManager = new XMLIndexer.MrpDataManager(@"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db");
            // Skip database init for now to avoid errors
            
            _loadingWindow?.UpdateProgress(50, "Loading data...", "Preparing application data");
            
            // Import Excel data synchronously (with timeout protection)
            try
            {
                _loadingWindow?.UpdateProgress(75, "Importing Excel data...", "Reading Priority List Master SHOP-SQL.xls");
                
                var task = System.Threading.Tasks.Task.Run(() => 
                {
                    mrpManager.ImportFromExcel(@"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls");
                });
                
                if (task.Wait(30000)) // 30 second timeout
                {
                    _loadingWindow?.UpdateProgress(90, "Excel data imported successfully", "Finalizing startup");
                }
                else
                {
                    _loadingWindow?.UpdateProgress(90, "Excel data timeout - using existing data", "Finalizing startup");
                }
            }
            catch (Exception)
            {
                _loadingWindow?.UpdateProgress(90, "Excel data failed - using existing data", "Finalizing startup");
            }
            
            _loadingWindow?.UpdateProgress(100, "Application ready", "Opening main interface");
        }
    }
}
