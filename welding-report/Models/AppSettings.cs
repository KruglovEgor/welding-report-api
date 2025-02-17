namespace welding_report.Models
{
    public class AppSettings
    {
        public string TemplatePath { get; set; }
        public string UploadsFolder { get; set; }
        public string ReportStoragePath { get; set; }
        public List<string> AllowedEmailDomains { get; set; }
        public int MaxRowHeight { get; set; }
        public string WorksheetName { get; set; }

        public int MaxPhotoColumnWidth { get; set; }
    }
}
