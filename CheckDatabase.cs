using System;
using Microsoft.Data.Sqlite;

class CheckDatabase
{
    static void Main(string[] args)
    {
        string dbPath = @"c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
        string connectionString = $"Data Source={dbPath}";
        
        try
        {
            using var connection = new SqliteConnection(connectionString);
            connection.Open();
            
            // Check SalesOrders table
            using var cmd1 = new SqliteCommand("SELECT COUNT(*) FROM SalesOrders", connection);
            var salesOrderCount = Convert.ToInt32(cmd1.ExecuteScalar());
            Console.WriteLine($"Sales Orders in database: {salesOrderCount}");
            
            // Check ProgrammedParts table
            using var cmd2 = new SqliteCommand("SELECT COUNT(*) FROM ProgrammedParts", connection);
            var programmedPartsCount = Convert.ToInt32(cmd2.ExecuteScalar());
            Console.WriteLine($"Programmed Parts in database: {programmedPartsCount}");
            
            // Show sample sales orders
            if (salesOrderCount > 0)
            {
                Console.WriteLine("\nSample Sales Orders:");
                using var cmd3 = new SqliteCommand("SELECT SalesOrder FROM SalesOrders LIMIT 10", connection);
                using var reader = cmd3.ExecuteReader();
                while (reader.Read())
                {
                    Console.WriteLine($"  - {reader["SalesOrder"]}");
                }
            }
            
            // Show sample programmed parts
            if (programmedPartsCount > 0)
            {
                Console.WriteLine("\nSample Programmed Parts:");
                using var cmd4 = new SqliteCommand("SELECT PartNumber FROM ProgrammedParts LIMIT 10", connection);
                using var reader = cmd4.ExecuteReader();
                while (reader.Read())
                {
                    Console.WriteLine($"  - {reader["PartNumber"]}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}