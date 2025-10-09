using System;
using System.IO;
using System.Text;

namespace XMLIndexer
{
    /// <summary>
    /// Service for exporting cutlists to various formats (Excel, CSV, HTML)
    /// </summary>
    public class CutlistExporter
    {
        /// <summary>
        /// Export cutlist to HTML format for printing
        /// </summary>
        public string ExportToHtml(Cutlist cutlist)
        {
            var html = new StringBuilder();
            
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset='utf-8'>");
            html.AppendLine($"    <title>{cutlist.Title}</title>");
            html.AppendLine("    <style>");
            html.AppendLine("        body { font-family: Arial, sans-serif; margin: 20px; }");
            html.AppendLine("        .header { text-align: center; margin-bottom: 20px; }");
            html.AppendLine("        .header h1 { font-size: 24px; font-weight: bold; margin: 5px 0; }");
            html.AppendLine("        .header .gstn { font-size: 16px; margin: 5px 0; }");
            html.AppendLine("        .header .completion { font-size: 14px; font-weight: bold; margin: 10px 0; }");
            html.AppendLine("        table { width: 100%; border-collapse: collapse; margin-top: 10px; }");
            html.AppendLine("        th, td { border: 1px solid #000; padding: 8px; text-align: center; }");
            html.AppendLine("        th { background-color: #666; color: white; font-weight: bold; }");
            html.AppendLine("        .part-desc { text-align: left; }");
            html.AppendLine("        .checkbox { width: 20px; height: 20px; }");
            html.AppendLine("        @media print { body { margin: 10px; } }");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            
            // Header section
            html.AppendLine("    <div class='header'>");
            html.AppendLine($"        <h1>{cutlist.Title}</h1>");
            if (!string.IsNullOrEmpty(cutlist.GSTNNumber))
            {
                html.AppendLine($"        <div class='gstn'>GSTN-{cutlist.GSTNNumber}</div>");
            }
            html.AppendLine("        <div class='completion'>INITIAL UPON COMPLETION</div>");
            html.AppendLine("    </div>");
            
            // Data table
            html.AppendLine("    <table>");
            html.AppendLine("        <thead>");
            html.AppendLine("            <tr>");
            html.AppendLine("                <th>PROG</th>");
            html.AppendLine("                <th>PART DESCRIPTION</th>");
            html.AppendLine("                <th>QTY</th>");
            html.AppendLine("                <th>YQTY</th>");
            html.AppendLine("                <th>XAX</th>");
            html.AppendLine("                <th>YAX</th>");
            html.AppendLine("                <th>GA</th>");
            html.AppendLine("                <th>QUALITY</th>");
            html.AppendLine("                <th>ENG</th>");
            html.AppendLine("                <th>NST</th>");
            html.AppendLine("                <th>LSR</th>");
            html.AppendLine("                <th>PCH</th>");
            html.AppendLine("                <th>FRM</th>");
            html.AppendLine("                <th>PEM</th>");
            html.AppendLine("            </tr>");
            html.AppendLine("        </thead>");
            html.AppendLine("        <tbody>");
            
            foreach (var item in cutlist.Items)
            {
                html.AppendLine("            <tr>");
                html.AppendLine($"                <td>{item.PROG}</td>");
                html.AppendLine($"                <td class='part-desc'>{item.PartDescription}</td>");
                html.AppendLine($"                <td>{item.QTY}</td>");
                html.AppendLine($"                <td>{item.YQTY}</td>");
                html.AppendLine($"                <td>{item.XAX:F2}</td>");
                html.AppendLine($"                <td>{item.YAX:F2}</td>");
                html.AppendLine($"                <td>{item.GA}</td>");
                html.AppendLine($"                <td>{item.Quality}</td>");
                html.AppendLine($"                <td><input type='checkbox' class='checkbox' {(item.ENGCompleted ? "checked" : "")} /></td>");
                html.AppendLine($"                <td><input type='checkbox' class='checkbox' {(item.NSTCompleted ? "checked" : "")} /></td>");
                html.AppendLine($"                <td><input type='checkbox' class='checkbox' {(item.LSRCompleted ? "checked" : "")} /></td>");
                html.AppendLine($"                <td><input type='checkbox' class='checkbox' {(item.PCHCompleted ? "checked" : "")} /></td>");
                html.AppendLine($"                <td><input type='checkbox' class='checkbox' {(item.FRMCompleted ? "checked" : "")} /></td>");
                html.AppendLine($"                <td><input type='checkbox' class='checkbox' {(item.PEMCompleted ? "checked" : "")} /></td>");
                html.AppendLine("            </tr>");
            }
            
            html.AppendLine("        </tbody>");
            html.AppendLine("    </table>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");
            
            return html.ToString();
        }

        /// <summary>
        /// Export cutlist to CSV format
        /// </summary>
        public string ExportToCsv(Cutlist cutlist)
        {
            var csv = new StringBuilder();
            
            // Add header information
            csv.AppendLine($"# {cutlist.Title}");
            if (!string.IsNullOrEmpty(cutlist.GSTNNumber))
            {
                csv.AppendLine($"# GSTN-{cutlist.GSTNNumber}");
            }
            csv.AppendLine("# INITIAL UPON COMPLETION");
            csv.AppendLine();
            
            // Add column headers
            csv.AppendLine("PROG,PART DESCRIPTION,QTY,YQTY,XAX,YAX,GA,QUALITY,ENG,NST,LSR,PCH,FRM,PEM");
            
            // Add data rows
            foreach (var item in cutlist.Items)
            {
                csv.AppendLine($"\"{item.PROG}\",\"{item.PartDescription}\",{item.QTY},{item.YQTY}," +
                              $"{item.XAX:F2},{item.YAX:F2},\"{item.GA}\",\"{item.Quality}\"," +
                              $"{(item.ENGCompleted ? "X" : "")},{(item.NSTCompleted ? "X" : "")}," +
                              $"{(item.LSRCompleted ? "X" : "")},{(item.PCHCompleted ? "X" : "")}," +
                              $"{(item.FRMCompleted ? "X" : "")},{(item.PEMCompleted ? "X" : "")}");
            }
            
            return csv.ToString();
        }

        /// <summary>
        /// Export cutlist to ERP XML format
        /// </summary>
        public string ExportToErp(Cutlist cutlist)
        {
            var xml = new StringBuilder();
            
            // XML header
            xml.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            xml.AppendLine("<ErpExchange>");
            xml.AppendLine("\t<Orders>");
            
            // Create ErpOrder for this job
            xml.AppendLine("\t\t<ErpOrder>");
            xml.AppendLine("\t\t\t<ImportType>NewOrder</ImportType>");
            xml.AppendLine($"\t\t\t<OrderNumber>{cutlist.JobNumber}</OrderNumber>");
            
            // Use current date as start date and target date
            var currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            xml.AppendLine($"\t\t\t<StartDate>{currentDate}</StartDate>");
            xml.AppendLine($"\t\t\t<TargetDate>{currentDate}</TargetDate>");
            
            xml.AppendLine("\t\t\t<ProductionStrategy>MaterialAdministrationOrder</ProductionStrategy>");
            xml.AppendLine("\t\t\t<Automatic>True</Automatic>");
            xml.AppendLine("\t\t\t<Parts>");
            
            // Add each part as ErpPart
            foreach (var item in cutlist.Items)
            {
                xml.AppendLine("\t\t\t\t<ErpPart>");
                xml.AppendLine($"\t\t\t\t\t<BysoftCode>{item.PROG}</BysoftCode>");
                xml.AppendLine($"\t\t\t\t\t<Debit>{item.QTY}</Debit>");
                xml.AppendLine($"\t\t\t\t\t<MaterialCode>{item.Quality}</MaterialCode>");
                xml.AppendLine("\t\t\t\t\t<Measure>Inch</Measure>");
                xml.AppendLine($"\t\t\t\t\t<Thickness>{item.GA}</Thickness>");
                xml.AppendLine("\t\t\t\t\t<RotationAllowance>Angle90</RotationAllowance>");
                xml.AppendLine("\t\t\t\t\t<FillPart>False</FillPart>");
                xml.AppendLine("\t\t\t\t</ErpPart>");
            }
            
            xml.AppendLine("\t\t\t</Parts>");
            xml.AppendLine("\t\t</ErpOrder>");
            xml.AppendLine("\t</Orders>");
            xml.AppendLine("</ErpExchange>");
            
            return xml.ToString();
        }

        /// <summary>
        /// Export burn list with multiple jobs to ERP XML format
        /// </summary>
        public string ExportBurnListToErp(List<BurnListJobItem> selectedJobs, MrpDataManager mrpManager)
        {
            var xml = new StringBuilder();
            
            // XML header
            xml.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            xml.AppendLine("<ErpExchange>");
            xml.AppendLine("\t<Orders>");
            
            var currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            
            // Create ErpOrder for each selected job
            foreach (var job in selectedJobs)
            {
                if (job.IsSelected && job.PartCount > 0)
                {
                    xml.AppendLine("\t\t<ErpOrder>");
                    xml.AppendLine("\t\t\t<ImportType>NewOrder</ImportType>");
                    xml.AppendLine($"\t\t\t<OrderNumber>{job.JobNumber}</OrderNumber>");
                    xml.AppendLine($"\t\t\t<StartDate>{currentDate}</StartDate>");
                    xml.AppendLine($"\t\t\t<TargetDate>{currentDate}</TargetDate>");
                    xml.AppendLine("\t\t\t<ProductionStrategy>MaterialAdministrationOrder</ProductionStrategy>");
                    xml.AppendLine("\t\t\t<Automatic>True</Automatic>");
                    xml.AppendLine("\t\t\t<Parts>");
                    
                    // Get actual component data for this job
                    var components = GetJobComponents(job.JobNumber, mrpManager);
                    
                    foreach (var component in components)
                    {
                        xml.AppendLine("\\t\\t\\t\\t<ErpPart>");
                        
                        // Strip file extension from PartNumber for BysoftCode
                        string bysoftCode = System.IO.Path.GetFileNameWithoutExtension(component.PartNumber);
                        xml.AppendLine($"\\t\\t\\t\\t\\t<BysoftCode>{bysoftCode}</BysoftCode>");
                        
                        xml.AppendLine($"\\t\\t\\t\\t\\t<Debit>{component.Quantity}</Debit>");
                        
                        // Use component material or map to standard codes
                        var materialCode = GetComponentMaterialCode(component.Material);
                        xml.AppendLine($"\\t\\t\\t\\t\\t<MaterialCode>{materialCode}</MaterialCode>");
                        
                        xml.AppendLine("\\t\\t\\t\\t\\t<Measure>Inch</Measure>");
                        
                        // Use component thickness or default
                        var thickness = component.Thickness > 0 ? component.Thickness.ToString("F4") : "0.1875";
                        xml.AppendLine($"\\t\\t\\t\\t\\t<Thickness>{thickness}</Thickness>");
                        
                        xml.AppendLine("\\t\\t\\t\\t\\t<RotationAllowance>Angle90</RotationAllowance>");
                        xml.AppendLine("\\t\\t\\t\\t\\t<FillPart>False</FillPart>");
                        xml.AppendLine("\\t\\t\\t\\t</ErpPart>");
                    }
                    
                    xml.AppendLine("\t\t\t</Parts>");
                    xml.AppendLine("\t\t</ErpOrder>");
                }
            }
            
            xml.AppendLine("\t</Orders>");
            xml.AppendLine("</ErpExchange>");
            
            return xml.ToString();
        }
        /// <summary>
        /// Get job components from database (similar to CutlistGenerator)
        /// </summary>
        private List<ComponentData> GetJobComponents(string jobNumber, MrpDataManager mrpManager)
        {
            var components = new List<ComponentData>();
            
            using var connection = new Microsoft.Data.Sqlite.SqliteConnection(mrpManager.ConnectionString);
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
        /// ComponentData class to hold component information
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
        }        /// <summary>
        /// Map component material to standard ERP material codes
        /// </summary>
        private string GetComponentMaterialCode(string componentMaterial)
        {
            if (string.IsNullOrEmpty(componentMaterial))
                return "MS"; // Default to mild steel
            
            // Map component materials to ERP codes
            string material = componentMaterial.ToUpper();
            
            if (material.Contains("A36") || material.Contains("STEEL") || material.Contains("CARBON"))
                return "A36 PLATE";
            if (material.Contains("GALV") || material.Contains("GALVANEAL"))
                return "GALVANEAL";
            if (material.Contains("MS") || material.Contains("MILD"))
                return "MS";
            if (material.Contains("ALUM") || material.Contains("ALUMINUM"))
                return "ALUMINUM";
            if (material.Contains("STAINLESS") || material.Contains("SS"))
                return "STAINLESS";
            
            // Default fallback
            return "MS";
        }


        
        /// <summary>
        /// Get material code from MaterialTable or use default mapping
        /// </summary>
        private string GetMaterialCode(string partNumber, MrpDataManager mrpManager)
        {
            // Try to get material from MaterialTable first
            var materialInfo = mrpManager.GetMaterialInfo(partNumber);
            if (!string.IsNullOrEmpty(materialInfo))
            {
                // Map common material types to ERP codes
                if (materialInfo.Contains("A36")) return "A36 PLATE";
                if (materialInfo.Contains("GALV")) return "GALVANEAL";
                if (materialInfo.Contains("MS") || materialInfo.Contains("MILD")) return "MS";
                return materialInfo.ToUpper();
            }
            
            // Default fallback
            return "MS";
        }
        
        /// <summary>
        /// Get thickness from MaterialTable or estimate from part number
        /// </summary>
        private string GetThickness(string partNumber, MrpDataManager mrpManager)
        {
            // Try to get thickness from MaterialTable first
            var thickness = mrpManager.GetThickness(partNumber);
            if (thickness > 0)
            {
                return thickness.ToString("F4");
            }
            
            // Default fallback
            return "0.1875";
        }

        /// <summary>
        /// Save cutlist to file
        /// </summary>
        public void SaveToFile(Cutlist cutlist, string filePath, CutlistFormat format = CutlistFormat.Html)
        {
            string content = format switch
            {
                CutlistFormat.Html => ExportToHtml(cutlist),
                CutlistFormat.Csv => ExportToCsv(cutlist),
                CutlistFormat.Erp => ExportToErp(cutlist),
                _ => throw new ArgumentException("Unsupported format", nameof(format))
            };

            File.WriteAllText(filePath, content, Encoding.UTF8);
        }

        /// <summary>
        /// Generate a default filename for the cutlist
        /// </summary>
        public string GenerateFileName(Cutlist cutlist, CutlistFormat format)
        {
            var extension = format switch
            {
                CutlistFormat.Html => ".html",
                CutlistFormat.Csv => ".csv", 
                CutlistFormat.Erp => ".erp",
                _ => ".txt"
            };

            var safeJobNumber = cutlist.JobNumber.Replace("-", "_").Replace(" ", "_");
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            
            return $"Cutlist_{safeJobNumber}_{timestamp}{extension}";
        }
    }

    /// <summary>
    /// Supported cutlist export formats
    /// </summary>
    public enum CutlistFormat
    {
        Html,
        Csv,
        Erp
    }
}







