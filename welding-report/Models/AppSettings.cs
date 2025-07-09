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
        public string ProjectReportRowColor1 { get; set; } 
        public string ProjectReportRowColor2 { get; set; }
        public int MaxPhotoWidthPx { get; set; }
        public int MaxPhotoHeightPx { get; set; }
        public int PhotoJpegQuality { get; set; }
        public string WeldingPhotoCachePath { get; set; }
    }
}
