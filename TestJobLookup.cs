using System;
using System.Data.SQLite;

class TestJobLookup
{
    static void Main()
    {
        string dbPath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
        string jobNumber = "IK3NC-0000";
        
        Console.WriteLine($"Testing job lookup for: {jobNumber}");
        Console.WriteLine($"Database path: {dbPath}");
        
        using var connection = new SQLiteConnection($"Data Source={dbPath}");
        connection.Open();
        
        // Test our JOIN query
        string query = @"
        SELECT m.*, 
               COALESCE(x.XmlStatus, 'No XML') as XmlStatus,
               COALESCE(x.HighestRelease, 0) as HighestRelease,
               COALESCE(x.ComponentCount, 0) as ComponentCount
        FROM MrpPriorityList m
        LEFT JOIN XMLIndex x ON m.PartNumber = x.PartNumber AND m.Revision = x.Revision
        WHERE m.JobNumber LIKE @jobNumber";
        
        using var command = new SQLiteCommand(query, connection);
        command.Parameters.AddWithValue("@jobNumber", $"%{jobNumber}%");
        
        Console.WriteLine("\nExecuting query...");
        using var reader = command.ExecuteReader();
        
        bool found = false;
        while (reader.Read())
        {
            found = true;
            Console.WriteLine($"Found job: {reader["JobNumber"]}");
            Console.WriteLine($"Part Number: {reader["PartNumber"]}");
            Console.WriteLine($"Revision: {reader["Revision"]}");
            Console.WriteLine($"Description: {reader["Description"]}");
            Console.WriteLine($"XML Status: {reader["XmlStatus"]}");
            Console.WriteLine($"Highest Release: {reader["HighestRelease"]}");
            Console.WriteLine($"Component Count: {reader["ComponentCount"]}");
            Console.WriteLine("---");
        }
        
        if (!found)
        {
            Console.WriteLine("No jobs found matching the search criteria.");
            
            // Check if data exists at all
            using var countCommand = new SQLiteCommand("SELECT COUNT(*) FROM MrpPriorityList", connection);
            var count = countCommand.ExecuteScalar();
            Console.WriteLine($"Total records in MrpPriorityList: {count}");
        }
        
        connection.Close();
        Console.WriteLine("Test completed.");
    }
}