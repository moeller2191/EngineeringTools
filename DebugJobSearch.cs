using System;
using System.Collections.Generic;
using XMLIndexer;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== Debug Job Search ===");
        
        string dbPath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
        Console.WriteLine($"Database path: {dbPath}");
        
        try
        {
            var mrpManager = new MrpDataManager(dbPath);
            
            // Test 1: Search for exact job
            Console.WriteLine("\n1. Searching for exact job 'IK3NC-0000':");
            var exactResults = mrpManager.GetMrpDataForJob("IK3NC-0000");
            Console.WriteLine($"   Found {exactResults.Count} results");
            foreach (var item in exactResults)
            {
                Console.WriteLine($"   - {item.JobNumber}: {item.PartNumber} ({item.Description})");
            }
            
            // Test 2: Search for partial job
            Console.WriteLine("\n2. Searching for partial job 'IK3NC':");
            var partialResults = mrpManager.GetMrpDataForJob("IK3NC");
            Console.WriteLine($"   Found {partialResults.Count} results");
            foreach (var item in partialResults)
            {
                Console.WriteLine($"   - {item.JobNumber}: {item.PartNumber} ({item.Description})");
            }
            
            // Test 3: Get all MRP data
            Console.WriteLine("\n3. Getting all MRP data:");
            var allData = mrpManager.GetAllMrpData();
            Console.WriteLine($"   Total jobs in database: {allData.Count}");
            
            // Show first few jobs
            Console.WriteLine("   First 10 jobs:");
            for (int i = 0; i < Math.Min(10, allData.Count); i++)
            {
                var item = allData[i];
                Console.WriteLine($"   {i+1}. {item.JobNumber}: {item.PartNumber}");
            }
            
            // Test 4: Search for any job containing '0000'
            Console.WriteLine("\n4. Searching for jobs containing '0000':");
            var zeroResults = mrpManager.GetMrpDataForJob("0000");
            Console.WriteLine($"   Found {zeroResults.Count} results");
            foreach (var item in zeroResults)
            {
                Console.WriteLine($"   - {item.JobNumber}: {item.PartNumber}");
            }
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}