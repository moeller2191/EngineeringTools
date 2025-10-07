using System;
using System.IO;

class Program
{
    static void Main()
    {
        string excelPath = @"C:\Scripts\EngineeringTools\Priority List Master SHOP-SQL.xls";
        
        try
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                Console.WriteLine("Excel COM not available");
                return;
            }
            
            dynamic excelApp = Activator.CreateInstance(excelType);
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            
            dynamic workbook = excelApp.Workbooks.Open(excelPath);
            dynamic worksheet = workbook.Worksheets["Priority List"];
            
            dynamic usedRange = worksheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;
            
            Console.WriteLine($"=== Excel Analysis ===");
            Console.WriteLine($"Total columns: {lastCol}");
            Console.WriteLine($"Total rows: {lastRow}");
            Console.WriteLine();
            
            Console.WriteLine("=== Column Headers (First 30) ===");
            for (int col = 1; col <= Math.Min(lastCol, 30); col++)
            {
                dynamic headerCell = worksheet.Cells[1, col];
                string header = headerCell.Value2?.ToString()?.Trim() ?? "NULL";
                Console.WriteLine($"Column {col}: '{header}'");
            }
            
            Console.WriteLine();
            Console.WriteLine("=== Sample Row 2 Data ===");
            for (int col = 1; col <= Math.Min(lastCol, 30); col++)
            {
                dynamic dataCell = worksheet.Cells[2, col];
                string value = dataCell.Value2?.ToString()?.Trim() ?? "NULL";
                Console.WriteLine($"Col {col}: '{value}'");
            }
            
            workbook.Close(false);
            excelApp.Quit();
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        
        Console.WriteLine("Press any key to continue...");
        Console.ReadKey();
    }
}