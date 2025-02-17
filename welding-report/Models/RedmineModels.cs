using System.Text.Json.Serialization;
using System.Text.Json;

namespace welding_report.Models
{
    public class RedmineReportData
    {
        public string ReportNumber { get; set; }
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
        public string JointNumbers { get; set; }
        public List<string> PhotoUrls { get; set; } = new();
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
