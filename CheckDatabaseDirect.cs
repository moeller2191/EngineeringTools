using System;
using Microsoft.Data.Sqlite;

class Program
{
    static void Main()
    {
        string dbPath = @"c:\Scripts\EngineeringTools\XMLIndexer\EngDataImport.db";
        string connectionString = $"Data Source={dbPath}";
        
        Console.WriteLine($"Checking database: {dbPath}");
        Console.WriteLine($"Database exists: {System.IO.File.Exists(dbPath)}");
        Console.WriteLine();
        
        try
        {
            using var connection = new SqliteConnection(connectionString);
            connection.Open();
            
            // Check SalesOrders table
            using var soCmd = new SqliteCommand("SELECT COUNT(*) FROM SalesOrders", connection);
            var salesOrderCount = Convert.ToInt32(soCmd.ExecuteScalar());
            Console.WriteLine($"SalesOrders count: {salesOrderCount}");
            
            if (salesOrderCount > 0 && salesOrderCount <= 10)
            {
                using var soSampleCmd = new SqliteCommand("SELECT SalesOrder FROM SalesOrders LIMIT 10", connection);
                using var soReader = soSampleCmd.ExecuteReader();
                Console.WriteLine("Sample sales orders:");
                while (soReader.Read())
                {
                    Console.WriteLine($"  - {soReader.GetString(0)}");
                }
            }
            Console.WriteLine();
            
            // Check ProgrammedParts table
            using var ppCmd = new SqliteCommand("SELECT COUNT(*) FROM ProgrammedParts", connection);
            var programmedPartsCount = Convert.ToInt32(ppCmd.ExecuteScalar());
            Console.WriteLine($"ProgrammedParts count: {programmedPartsCount}");
            
            if (programmedPartsCount > 0 && programmedPartsCount <= 10)
            {
                using var ppSampleCmd = new SqliteCommand("SELECT PartNumber FROM ProgrammedParts LIMIT 10", connection);
                using var ppReader = ppSampleCmd.ExecuteReader();
                Console.WriteLine("Sample programmed parts:");
                while (ppReader.Read())
                {
                    Console.WriteLine($"  - {ppReader.GetString(0)}");
                }
            }
            Console.WriteLine();
            
            // Check MrpPriorityList table
            using var mrpCmd = new SqliteCommand("SELECT COUNT(*) FROM MrpPriorityList", connection);
            var mrpCount = Convert.ToInt32(mrpCmd.ExecuteScalar());
            Console.WriteLine($"MrpPriorityList count: {mrpCount}");
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        
        Console.WriteLine("\nPress any key to continue...");
        Console.ReadKey();
    }
}