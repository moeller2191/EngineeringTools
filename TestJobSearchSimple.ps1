# Simple test for job search
Write-Host "Testing Job Search for IK3NC-0000" -ForegroundColor Green

$databasePath = "C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db"
$excelPath = "C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls"

Write-Host "1. Checking if database exists..."
if (Test-Path $databasePath) {
    Write-Host "   ✓ Database exists" -ForegroundColor Green
} else {
    Write-Host "   ✗ Database not found" -ForegroundColor Red
    exit 1
}

Write-Host "2. Checking if Excel file exists..."
if (Test-Path $excelPath) {
    Write-Host "   ✓ Excel file exists" -ForegroundColor Green
} else {
    Write-Host "   ✗ Excel file not found" -ForegroundColor Red
    exit 1
}

Write-Host "3. Building test console app..."
$testCode = @"
using System;
using XMLIndexer;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var mrpManager = new XMLIndexer.MrpDataManager(@"$databasePath");
                
                Console.WriteLine("Importing Excel data...");
                bool success = mrpManager.ImportFromExcel(@"$excelPath");
                Console.WriteLine($"Import success: {success}");
                
                Console.WriteLine("Searching for IK3NC-0000...");
                var results = mrpManager.GetMrpDataForJob("IK3NC-0000");
                Console.WriteLine($"Found {results.Count} results");
                
                foreach (var item in results)
                {
                    Console.WriteLine($"Job: {item.JobNumber}, Part: {item.PartNumber}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
        }
    }
}
"@

$testCode | Out-File -FilePath "TestJobSearch.cs" -Encoding UTF8

Write-Host "4. Compiling and running test..."
dotnet run TestJobSearch.cs --project XMLIndexer