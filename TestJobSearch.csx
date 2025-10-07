using System;
using System.Collections.Generic;
using System.Linq;
using XMLIndexer;

var databasePath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
var excelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";

Console.WriteLine("=== Job Search Debug Test ===");

try
{
    var mrpManager = new MrpDataManager(databasePath);
    
    Console.WriteLine("1. Importing Excel data...");
    bool success = mrpManager.ImportFromExcel(excelPath);
    Console.WriteLine($"   Import result: {success}");
    
    if (!success)
    {
        Console.WriteLine("   Excel import failed - job won't be found");
        return;
    }
    
    Console.WriteLine("2. Searching for 'IK3NC-0000'...");
    var results = mrpManager.GetMrpDataForJob("IK3NC-0000");
    Console.WriteLine($"   Found {results.Count} exact results");
    
    if (results.Count == 0)
    {
        Console.WriteLine("3. Trying partial search for 'IK3NC'...");
        var partialResults = mrpManager.GetMrpDataForJob("IK3NC");
        Console.WriteLine($"   Found {partialResults.Count} partial results");
        
        if (partialResults.Count > 0)
        {
            Console.WriteLine("   First few matches:");
            foreach (var item in partialResults.Take(5))
            {
                Console.WriteLine($"     Job: '{item.JobNumber}', Part: '{item.PartNumber}'");
            }
        }
    }
    else
    {
        Console.WriteLine("   SUCCESS! Found exact matches:");
        foreach (var item in results)
        {
            Console.WriteLine($"     Job: '{item.JobNumber}', Part: '{item.PartNumber}', Desc: '{item.Description}'");
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine($"ERROR: {ex.Message}");
    Console.WriteLine($"Stack: {ex.StackTrace}");
}

Console.WriteLine("Press any key to exit...");
Console.ReadKey();