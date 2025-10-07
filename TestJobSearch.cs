using System;
using XMLIndexer;

namespace EngineeringTools
{
    class TestJobSearch
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Testing Job Search...");
            
            string databasePath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
            string excelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
            
            try
            {
                // Test 1: Create MrpDataManager
                Console.WriteLine("1. Creating MrpDataManager...");
                var mrpManager = new MrpDataManager(databasePath);
                Console.WriteLine("   ✓ MrpDataManager created");
                
                // Test 2: Import Excel data
                Console.WriteLine("2. Importing Excel data...");
                bool importSuccess = mrpManager.ImportFromExcel(excelPath);
                Console.WriteLine($"   Import result: {importSuccess}");
                
                // Test 3: Search for IK3NC-0000
                Console.WriteLine("3. Searching for IK3NC-0000...");
                var results = mrpManager.GetMrpDataForJob("IK3NC-0000");
                Console.WriteLine($"   Found {results.Count} results");
                
                foreach (var item in results)
                {
                    Console.WriteLine($"   - Job: {item.JobNumber}, Part: {item.PartNumber}, Desc: {item.Description}");
                }
                
                // Test 4: Search for IK3NC (partial)
                Console.WriteLine("4. Searching for IK3NC (partial)...");
                var partialResults = mrpManager.GetMrpDataForJob("IK3NC");
                Console.WriteLine($"   Found {partialResults.Count} results");
                
                foreach (var item in partialResults.Take(5)) // Show first 5
                {
                    Console.WriteLine($"   - Job: {item.JobNumber}, Part: {item.PartNumber}");
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}