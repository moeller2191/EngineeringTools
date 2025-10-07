using System;
using Microsoft.Data.Sqlite;

namespace XMLIndexer
{
    class DatabaseDebugger
    {
        static void Main(string[] args)
        {
            string dbPath = @"c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
            string connectionString = $"Data Source={dbPath}";
            
            Console.WriteLine($"Checking XMLIndex database at: {dbPath}");
            
            try
            {
                using var connection = new SqliteConnection(connectionString);
                connection.Open();
                
                Console.WriteLine("\n=== DATABASE TABLES ===");
                using var cmd1 = new SqliteCommand("SELECT name FROM sqlite_master WHERE type='table'", connection);
                using var reader1 = cmd1.ExecuteReader();
                while (reader1.Read())
                {
                    Console.WriteLine($"Table: {reader1["name"]}");
                }
                reader1.Close();
                
                Console.WriteLine("\n=== XMLFILES COUNT ===");
                using var cmd2 = new SqliteCommand("SELECT COUNT(*) as count FROM XMLFiles", connection);
                var count = cmd2.ExecuteScalar();
                Console.WriteLine($"XMLFiles count: {count}");
                
                Console.WriteLine("\n=== SAMPLE XMLFILES ===");
                using var cmd3 = new SqliteCommand("SELECT PartNumber, Revision, Release, FileName FROM XMLFiles LIMIT 10", connection);
                using var reader3 = cmd3.ExecuteReader();
                while (reader3.Read())
                {
                    Console.WriteLine($"Part: {reader3["PartNumber"]}, Rev: {reader3["Revision"]}, Rel: {reader3["Release"]}, File: {reader3["FileName"]}");
                }
                reader3.Close();
                
                Console.WriteLine("\n=== COMPONENTS COUNT ===");
                using var cmd4 = new SqliteCommand("SELECT COUNT(*) as count FROM Components", connection);
                var compCount = cmd4.ExecuteScalar();
                Console.WriteLine($"Components count: {compCount}");
                
                Console.WriteLine("\n=== MRP TABLE CHECK ===");
                try
                {
                    using var cmd5 = new SqliteCommand("SELECT COUNT(*) as count FROM MrpPriorityList", connection);
                    var mrpCount = cmd5.ExecuteScalar();
                    Console.WriteLine($"MrpPriorityList count: {mrpCount}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MrpPriorityList table error: {ex.Message}");
                }
                
                Console.WriteLine("\n=== SEARCH FOR SAMPLE PART (I-02250170) ===");
                using var cmd6 = new SqliteCommand("SELECT PartNumber, Revision, Release FROM XMLFiles WHERE PartNumber LIKE '%02250170%' LIMIT 5", connection);
                using var reader6 = cmd6.ExecuteReader();
                while (reader6.Read())
                {
                    Console.WriteLine($"Found: {reader6["PartNumber"]}, Rev: {reader6["Revision"]}, Rel: {reader6["Release"]}");
                }
                reader6.Close();
                
                connection.Close();
                Console.WriteLine("\nDatabase check complete!");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}