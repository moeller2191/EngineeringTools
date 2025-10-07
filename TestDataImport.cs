using XMLIndexer;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Testing Excel Data Import...");
        Console.WriteLine("================================");
        
        try
        {
            var databasePath = @"C:\Scripts\EngineeringTools\XMLIndexer\XMLIndex.db";
            var mrpManager = new MrpDataManager(databasePath);
            
            Console.WriteLine("Starting ImportRealData...");
            mrpManager.ImportRealData();
            Console.WriteLine("Import completed!");
            
            // Test some lookups
            Console.WriteLine("\nTesting lookups...");
            
            // Test sales order check
            var testSalesOrder = "41971";
            bool found = mrpManager.CheckSalesOrderInDatabase(testSalesOrder);
            Console.WriteLine($"Sales Order {testSalesOrder} found: {found}");
            
            // Test programmed part check
            var testPart = "A-90066706";
            bool programmed = mrpManager.CheckPartProgrammedInDatabase(testPart);
            Console.WriteLine($"Part {testPart} programmed: {programmed}");
            
            // Show counts
            var allSalesOrders = mrpManager.GetCheckedSalesOrdersFromDatabase();
            var allProgrammedParts = mrpManager.GetProgrammedPartsFromDatabase();
            
            Console.WriteLine($"\nTotal Sales Orders: {allSalesOrders.Count}");
            Console.WriteLine($"Total Programmed Parts: {allProgrammedParts.Count}");
            
            if (allSalesOrders.Count > 0)
            {
                Console.WriteLine($"First few sales orders: {string.Join(", ", allSalesOrders.Take(5))}");
            }
            
            if (allProgrammedParts.Count > 0)
            {
                Console.WriteLine($"First few programmed parts: {string.Join(", ", allProgrammedParts.Take(5))}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}