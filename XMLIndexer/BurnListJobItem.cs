namespace XMLIndexer
{
    /// <summary>
    /// Represents a job item in the burn list
    /// </summary>
    public class BurnListJobItem
    {
        public string JobNumber { get; set; } = "";
        public string Status { get; set; } = "";
        public int PartCount { get; set; }
        public string XmlStatus { get; set; } = "";
        public string Description { get; set; } = "";
        public bool IsSelected { get; set; } = true;
    }
}
