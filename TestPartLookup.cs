using System;
using Microsoft.Data.Sqlite;

Console.WriteLine("=== MRP vs XML Part Number Debug ===");

string xmlDbPath = @"c:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
string xmlConnectionString = $"Data Source={xmlDbPath}";

Console.WriteLine($"Checking XML database: {xmlDbPath}");

try
{
    using var xmlConnection = new SqliteConnection(xmlConnectionString);
    xmlConnection.Open();
    
    Console.WriteLine("\n=== XML DATABASE SAMPLE PARTS ===");
    using var xmlCmd = new SqliteCommand("SELECT PartNumber, Revision, Release FROM XMLFiles WHERE PartNumber IS NOT NULL AND PartNumber != '' ORDER BY PartNumber LIMIT 20", xmlConnection);
    using var xmlReader = xmlCmd.ExecuteReader();
    while (xmlReader.Read())
    {
        Console.WriteLine($"XML Part: '{xmlReader["PartNumber"]}', Rev: '{xmlReader["Revision"]}', Rel: '{xmlReader["Release"]}'");
    }
    xmlReader.Close();
    
    Console.WriteLine("\n=== SEARCH TESTS ===");
    string[] testParts = { "I-02250170", "02250170", "SULL-I-02250170", "H123456", "J123456" };
    
    foreach (string testPart in testParts)
    {
        Console.WriteLine($"\nSearching for: '{testPart}'");
        using var searchCmd = new SqliteCommand("SELECT PartNumber, Revision, Release FROM XMLFiles WHERE PartNumber LIKE @search LIMIT 5", xmlConnection);
        searchCmd.Parameters.AddWithValue("@search", $"%{testPart}%");
        using var searchReader = searchCmd.ExecuteReader();
        int count = 0;
        while (searchReader.Read())
        {
            Console.WriteLine($"  Found: '{searchReader["PartNumber"]}', Rev: '{searchReader["Revision"]}', Rel: '{searchReader["Release"]}'");
            count++;
        }
        if (count == 0)
        {
            Console.WriteLine($"  No matches found for '{testPart}'");
        }
        searchReader.Close();
    }
    
    xmlConnection.Close();
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}

Console.WriteLine("\nPress any key to exit...");
Console.ReadKey();