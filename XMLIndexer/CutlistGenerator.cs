using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace XMLIndexer
{
    /// <summary>
    /// Service for generating cutlists from job data and material specifications
    /// </summary>
    public class CutlistGenerator
    {
        private readonly string _connectionString;
        private readonly List<MaterialSpec> _materialSpecs;

        public CutlistGenerator(string connectionString)
        {
            _connectionString = connectionString;
            _materialSpecs = LoadMaterialSpecs();
        }

        /// <summary>
        /// Generate a cutlist for the specified job number
        /// </summary>
        public Cutlist GenerateCutlist(string jobNumber)
        {
            var cutlist = new Cutlist 
            { 
                JobNumber = jobNumber,
                GSTNNumber = ExtractGSTNFromJob(jobNumber)
            };

            // Get components for this job from XML data
            var components = GetJobComponents(jobNumber);
            
            // Group components by part number and aggregate quantities
            var groupedComponents = components
                .GroupBy(c => c.PartNumber)
                .Select(g => new 
                {
                    PartNumber = g.Key,
                    Description = g.First().Description,
                    TotalQuantity = g.Sum(c => c.Quantity),
                    Material = g.First().Material,
                    XDimension = g.First().XDimension,
                    YDimension = g.First().YDimension,
                    Thickness = g.First().Thickness
                });

            // Convert to cutlist items
            foreach (var component in groupedComponents)
            {
                var materialSpec = FindMaterialSpec(component.Material, component.Thickness);
                
                var item = new CutlistItem
                {
                    PROG = component.PartNumber,
                    PartDescription = component.Description,
                    QTY = component.TotalQuantity,
                    YQTY = 1, // Default to 1, may need business logic
                    XAX = component.XDimension,
                    YAX = component.YDimension,
                    GA = materialSpec?.Gauge ?? "Unknown",
                    Quality = materialSpec?.SolidWorksMaterialCode ?? "Unknown"
                };

                cutlist.Items.Add(item);
            }

            return cutlist;
        }

        /// <summary>
        /// Load material specifications from the database
        /// </summary>
        private List<MaterialSpec> LoadMaterialSpecs()
        {
            var specs = new List<MaterialSpec>();
            
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();

            // Check if MaterialTable exists, if not return empty list
            var tableCheckCmd = connection.CreateCommand();
            tableCheckCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='MaterialTable';";
            var tableExists = tableCheckCmd.ExecuteScalar() != null;
            
            if (!tableExists) return specs;

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT ID, MaterialPartNo, BysoftMaterialCode, SolidWorksMaterialCode, 
                       Thickness, Gauge, ScrapFactor, Pounds, SheetMajor, SheetMinor 
                FROM MaterialTable";

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                specs.Add(new MaterialSpec
                {
                    ID = reader.GetInt32(0),
                    MaterialPartNo = reader.GetString(1),
                    BysoftMaterialCode = reader.GetString(2),
                    SolidWorksMaterialCode = reader.GetString(3),
                    Thickness = reader.GetDouble(4),
                    Gauge = reader.GetString(5),
                    ScrapFactor = reader.GetDouble(6),
                    Pounds = reader.GetDouble(7),
                    SheetMajor = reader.IsDBNull(8) ? "" : reader.GetString(8),
                    SheetMinor = reader.IsDBNull(9) ? "" : reader.GetString(9)
                });
            }

            return specs;
        }

        /// <summary>
        /// Get component data for a specific job from XML analysis
        /// </summary>
        private List<ComponentData> GetJobComponents(string jobNumber)
        {
            var components = new List<ComponentData>();
            
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT DISTINCT c.PartNumber, c.Description, c.Quantity, c.Material, 
                       c.XDimension, c.YDimension, c.Thickness
                FROM Components c
                INNER JOIN XMLFiles x ON c.XMLFileID = x.ID
                WHERE x.FileName LIKE @jobPattern";

            command.Parameters.AddWithValue("@jobPattern", $"%{jobNumber}%");

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                components.Add(new ComponentData
                {
                    PartNumber = reader.IsDBNull(0) ? "" : reader.GetString(0),
                    Description = reader.IsDBNull(1) ? "" : reader.GetString(1),
                    Quantity = reader.IsDBNull(2) ? 1 : reader.GetInt32(2),
                    Material = reader.IsDBNull(3) ? "" : reader.GetString(3),
                    XDimension = reader.IsDBNull(4) ? 0.0 : reader.GetDouble(4),
                    YDimension = reader.IsDBNull(5) ? 0.0 : reader.GetDouble(5),
                    Thickness = reader.IsDBNull(6) ? 0.0 : reader.GetDouble(6)
                });
            }

            return components;
        }

        /// <summary>
        /// Find the best matching material specification for given material and thickness
        /// </summary>
        private MaterialSpec? FindMaterialSpec(string material, double thickness)
        {
            // First try exact material code match
            var exactMatch = _materialSpecs.FirstOrDefault(m => 
                m.BysoftMaterialCode.Equals(material, StringComparison.OrdinalIgnoreCase));
            
            if (exactMatch != null) return exactMatch;

            // Try material name matching (e.g., "Aluminum" matches "ALUM")
            var materialMatch = _materialSpecs.FirstOrDefault(m => 
                material.Contains(m.BysoftMaterialCode.Substring(0, Math.Min(4, m.BysoftMaterialCode.Length)), 
                StringComparison.OrdinalIgnoreCase));

            if (materialMatch != null) return materialMatch;

            // Try thickness-based matching as fallback
            var thicknessMatch = _materialSpecs
                .Where(m => Math.Abs(m.Thickness - thickness) < 0.001)
                .OrderBy(m => Math.Abs(m.Thickness - thickness))
                .FirstOrDefault();

            return thicknessMatch;
        }

        /// <summary>
        /// Extract GSTN number from job number (business logic may vary)
        /// </summary>
        private string ExtractGSTNFromJob(string jobNumber)
        {
            // Extract numeric part from job number (e.g., H1394-0000 -> 4190997)
            // This may need adjustment based on actual business rules
            var match = Regex.Match(jobNumber, @"(\d+)");
            return match.Success ? match.Groups[1].Value : "";
        }

        /// <summary>
        /// Helper class for component data
        /// </summary>
        private class ComponentData
        {
            public string PartNumber { get; set; } = "";
            public string Description { get; set; } = "";
            public int Quantity { get; set; }
            public string Material { get; set; } = "";
            public double XDimension { get; set; }
            public double YDimension { get; set; }
            public double Thickness { get; set; }
        }
    }
}