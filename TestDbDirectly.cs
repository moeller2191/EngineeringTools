using System;
using System.Data.SQLite;

public class TestDbDirectly
{
    public static void Main()
    {
        try
        {
            string connectionString = "Data Source=XMLIndex.db;Version=3;";
            Console.WriteLine($"Connecting to: {connectionString}");
            
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                Console.WriteLine("Database connection opened successfully.");
                
                // First, let's see all tables
                var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", connection);
                var reader = command.ExecuteReader();
                Console.WriteLine("\nTables in database:");
                while (reader.Read())
                {
                    Console.WriteLine($"  {reader["name"]}");
                }
                reader.Close();
                
                // Check if MrpPriorityList table exists and has data
                command = new SQLiteCommand("SELECT COUNT(*) FROM MrpPriorityList;", connection);
                var count = command.ExecuteScalar();
                Console.WriteLine($"\nTotal records in MrpPriorityList: {count}");
                
                // Search for IK3NC-0000 specifically
                command = new SQLiteCommand("SELECT * FROM MrpPriorityList WHERE JobNumber = 'IK3NC-0000';", connection);
                reader = command.ExecuteReader();
                Console.WriteLine("\nExact search for 'IK3NC-0000':");
                bool found = false;
                while (reader.Read())
                {
                    found = true;
                    Console.WriteLine($"JobNumber: {reader["JobNumber"]}, PartNumber: {reader["PartNumber"]}, Revision: {reader["Revision"]}");
                }
                if (!found) Console.WriteLine("  No exact matches found.");
                reader.Close();
                
                // Case-insensitive search 
                command = new SQLiteCommand("SELECT * FROM MrpPriorityList WHERE UPPER(JobNumber) = UPPER('IK3NC-0000');", connection);
                reader = command.ExecuteReader();
                Console.WriteLine("\nCase-insensitive search for 'IK3NC-0000':");
                found = false;
                while (reader.Read())
                {
                    found = true;
                    Console.WriteLine($"JobNumber: {reader["JobNumber"]}, PartNumber: {reader["PartNumber"]}, Revision: {reader["Revision"]}");
                }
                if (!found) Console.WriteLine("  No case-insensitive matches found.");
                reader.Close();
                
                // Partial search 
                command = new SQLiteCommand("SELECT * FROM MrpPriorityList WHERE JobNumber LIKE '%IK3NC%' LIMIT 10;", connection);
                reader = command.ExecuteReader();
                Console.WriteLine("\nPartial search for 'IK3NC' (first 10):");
                found = false;
                while (reader.Read())
                {
                    found = true;
                    Console.WriteLine($"JobNumber: {reader["JobNumber"]}, PartNumber: {reader["PartNumber"]}, Revision: {reader["Revision"]}");
                }
                if (!found) Console.WriteLine("  No partial matches found.");
                reader.Close();
                
                // Show some sample data
                command = new SQLiteCommand("SELECT JobNumber, PartNumber, Revision FROM MrpPriorityList LIMIT 5;", connection);
                reader = command.ExecuteReader();
                Console.WriteLine("\nFirst 5 records in MrpPriorityList:");
                while (reader.Read())
                {
                    Console.WriteLine($"JobNumber: {reader["JobNumber"]}, PartNumber: {reader["PartNumber"]}, Revision: {reader["Revision"]}");
                }
                reader.Close();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
}