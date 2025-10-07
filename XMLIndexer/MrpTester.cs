using System;
using System.Linq;
using Microsoft.Data.Sqlite;

namespace XMLIndexer
{
    public static class MrpTester
    {
        public static void TestMrpFunctionality(string databasePath)
        {
            Console.WriteLine("=== MRP FUNCTIONALITY TEST ===");
            
            try
            {
                var mrpManager = new MrpDataManager(databasePath);
                
                Console.WriteLine("1. Testing database connection...");
                var connectionString = $"Data Source={databasePath}";
                using var connection = new SqliteConnection(connectionString);
                connection.Open();
                
                // Check if MRP table exists
                using var checkTableCmd = new SqliteCommand(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='MrpPriorityList'", 
                    connection);
                var tableName = checkTableCmd.ExecuteScalar()?.ToString();
                
                if (tableName == "MrpPriorityList")
                {
                    Console.WriteLine("✅ MrpPriorityList table exists");
                    
                    // Count records
                    using var countCmd = new SqliteCommand("SELECT COUNT(*) FROM MrpPriorityList", connection);
                    var count = Convert.ToInt32(countCmd.ExecuteScalar());
                    Console.WriteLine($"✅ Found {count} MRP records");
                    
                    // Test view
                    using var viewCmd = new SqliteCommand("SELECT COUNT(*) FROM vw_MrpWithXmlStatus", connection);
                    var viewCount = Convert.ToInt32(viewCmd.ExecuteScalar());
                    Console.WriteLine($"✅ MRP view has {viewCount} records");
                    
                    // Test MrpDataManager
                    Console.WriteLine("\n2. Testing MrpDataManager...");
                    var activeJobs = mrpManager.GetActiveMrpJobs();
                    Console.WriteLine($"✅ Found {activeJobs.Count} active jobs");
                    
                    if (activeJobs.Count > 0)
                    {
                        var firstJob = activeJobs.First();
                        Console.WriteLine($"   Sample job: {firstJob.JobNumber} - {firstJob.PartNumber} - {firstJob.XmlStatus}");
                        
                        // Test specific job lookup
                        var jobData = mrpManager.GetMrpDataForJob(firstJob.JobNumber);
                        Console.WriteLine($"✅ Job lookup returned {jobData.Count} items");
                    }
                    
                    // Note: Job history now stored in Excel instead of SQLite
                    Console.WriteLine("\n3. Job history now managed in Excel file");
                    
                } else {
                    Console.WriteLine("❌ MrpPriorityList table not found - need to run MRP_Schema.sql");
                    return;
                }
                
                Console.WriteLine("\n✅ All MRP functionality tests passed!");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ MRP test failed: {ex.Message}");
                Console.WriteLine($"Details: {ex}");
            }
        }
    }
}