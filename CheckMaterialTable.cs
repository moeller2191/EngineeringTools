using Microsoft.Data.Sqlite;
using System;

class Program
{
    static void Main()
    {
        var connectionString = @"Data Source=C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
        
        try
        {
            using var connection = new SqliteConnection(connectionString);
            connection.Open();
            
            Console.WriteLine("=== DATABASE TABLES ===");
            var tablesCmd = connection.CreateCommand();
            tablesCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;";
            
            using var reader = tablesCmd.ExecuteReader();
            while (reader.Read())
            {
                var tableName = reader.GetString(0);
                Console.WriteLine($"Table: {tableName}");
            }
            reader.Close();
            
            // Check MaterialTable specifically
            Console.WriteLine("\n=== MATERIAL TABLE CHECK ===");
            var materialCountCmd = connection.CreateCommand();
            materialCountCmd.CommandText = "SELECT COUNT(*) FROM MaterialTable;";
            
            try
            {
                var count = materialCountCmd.ExecuteScalar();
                Console.WriteLine($"MaterialTable has {count} rows");
                
                // Show first few entries
                var sampleCmd = connection.CreateCommand();
                sampleCmd.CommandText = "SELECT MaterialPartNo, BysoftMaterialCode, Gauge FROM MaterialTable LIMIT 5;";
                
                using var sampleReader = sampleCmd.ExecuteReader();
                Console.WriteLine("\nSample MaterialTable entries:");
                while (sampleReader.Read())
                {
                    Console.WriteLine($"  {sampleReader.GetString(0)} | {sampleReader.GetString(1)} | {sampleReader.GetString(2)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"MaterialTable error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Database error: {ex.Message}");
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}