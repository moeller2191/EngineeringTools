using System;
using Microsoft.Data.Sqlite;

namespace TestJobQuery
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Data Source=C:\\Scripts\\EngineeringTools\\XMLIndexer\\XMLIndex.db";
            
            Console.WriteLine("Searching for jobs containing 'IK3NC':");
            
            using var connection = new SqliteConnection(connectionString);
            connection.Open();
            
            // First, let's see all job numbers that contain IK3NC
            using var command1 = new SqliteCommand(@"
                SELECT DISTINCT JobNumber 
                FROM vw_MrpWithXmlStatus 
                WHERE JobNumber LIKE '%IK3NC%' 
                ORDER BY JobNumber", connection);
                
            using var reader1 = command1.ExecuteReader();
            bool foundAny = false;
            while (reader1.Read())
            {
                foundAny = true;
                Console.WriteLine($"Found job: {reader1.GetString("JobNumber")}");
            }
            reader1.Close();
            
            if (!foundAny)
            {
                Console.WriteLine("No jobs found containing 'IK3NC'");
                
                // Let's see what job numbers actually exist
                Console.WriteLine("\nFirst 10 job numbers in database:");
                using var command2 = new SqliteCommand(@"
                    SELECT DISTINCT JobNumber 
                    FROM vw_MrpWithXmlStatus 
                    ORDER BY JobNumber 
                    LIMIT 10", connection);
                    
                using var reader2 = command2.ExecuteReader();
                while (reader2.Read())
                {
                    Console.WriteLine($"  {reader2.GetString("JobNumber")}");
                }
            }
            
            // Now test the exact query that the application uses
            Console.WriteLine($"\nTesting exact query for 'IK3NC-0000':");
            using var command3 = new SqliteCommand(@"
                SELECT * FROM vw_MrpWithXmlStatus 
                WHERE JobNumber LIKE @jobNumber
                ORDER BY Priority ASC", connection);
                
            command3.Parameters.AddWithValue("@jobNumber", "%IK3NC-0000%");
            
            using var reader3 = command3.ExecuteReader();
            int count = 0;
            while (reader3.Read())
            {
                count++;
                Console.WriteLine($"  Job: {reader3.GetString("JobNumber")}, Part: {reader3.GetString("PartNumber")}");
            }
            Console.WriteLine($"Total results: {count}");
        }
    }
}