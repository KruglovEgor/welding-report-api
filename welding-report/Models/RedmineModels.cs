using System.Text.Json.Serialization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace welding_report.Models
{
    public class RedmineReportData
    {
        public string ReportNumber { get; set; }
        public int JointsCount { get; set; }
        public List<JointGroup> Groups { get; set; } = new();
    }

    public class JointGroup
    {
        public int ActParagraph { get; set; }
        public string EquipmentType { get; set; }
        public string PipelineNumber { get; set; }
        public double DiameterMm { get; set; }
        public double DiameterInches { get; set; }
        
        public List<JointEntry> Entries { get; set; } = new();
    }

    public class JointEntry
    {
        public string Contractor { get; set; }
        public SortedDictionary<string, List<string>> JointPhotoMap { get; set; }
            = new SortedDictionary<string, List<string>>(new JointNumbersComparer());

    }

    public class JointNumbersComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            int firstX = ExtractFirstNumber(x);
            int firstY = ExtractFirstNumber(y);
            return firstX.CompareTo(firstY);
        }

        private static int ExtractFirstNumber(string input)
        {
            var match = Regex.Match(input, @"^\D*(\d+)");
            return match.Success ? int.Parse(match.Groups[1].Value) : int.MaxValue;
        }
    }

    public class RedmineIssueResponse
    {
        [JsonPropertyName("issue")]
        public RedmineIssue Issue { get; set; }
    }

    public class RedmineIssue
    {
        [JsonPropertyName("subject")]
        public string Subject { get; set; }

        [JsonPropertyName("custom_fields")]
        public List<RedmineCustomField> CustomFields { get; set; }
    }

    public class RedmineIssueListResponse
    {
        [JsonPropertyName("issues")]
        public List<RedmineChildIssue> Issues { get; set; }
    }

    public class RedmineChildIssue
    {
        [JsonPropertyName("custom_fields")]
        public List<RedmineCustomField> CustomFields { get; set; }

        [JsonPropertyName("attachments")]
        public List<RedmineAttachment> Attachments { get; set; }
    }

    public class RedmineCustomField
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("value")]
        public JsonElement Value { get; set; }
    }

    public class RedmineAttachment
    {
        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("content_url")]
        public string ContentUrl { get; set; }
    }
}
