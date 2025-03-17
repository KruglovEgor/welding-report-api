using System.Text.Json.Serialization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace welding_report.Models
{
    public class WeldingProjectReportData
    {
        public string Name { get; set; }
        public string Identifier { get; set; }
        public List<WeldingReportData> Acts { get; set; } = new();
    }

    public class WeldingProjectResponse
    {
        [JsonPropertyName("project")]
        public WeldingProject Project { get; set; }
    }

    public class WeldingProject
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

    }

    public class WeldingReportData
    {
        public string ReportNumber { get; set; }
        public int JointsCountFact { get; set; }
        public int JointsCountPlan { get; set; }
        public double DiametrInchesFact { get; set; }
        public double DiametrInchesPlan { get; set; }
        public List<JointGroup> Groups { get; set; } = new();
    }

    public class JointGroup
    {
        public int ActParagraph { get; set; }
        public string EquipmentType { get; set; }
        public string PipelineNumber { get; set; }
        public double DiameterMm { get; set; }
        public double DiameterInches { get; set; }
        public int JointsCount { get; set; }
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

    public class WeldingIssueResponse
    {
        [JsonPropertyName("issue")]
        public WeldingIssue Issue { get; set; }
    }

    public class WeldingIssue
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("subject")]
        public string Subject { get; set; }

        [JsonPropertyName("custom_fields")]
        public List<WeldingCustomField> CustomFields { get; set; }
    }

    public class WeldingIssueListResponse
    {
        [JsonPropertyName("issues")]
        public List<WeldingIssue> Issues { get; set; }
    }

    public class RedmineChildIssueListResponse
    {
        [JsonPropertyName("issues")]
        public List<RedmineChildIssue> Issues { get; set; }
    }

    public class RedmineChildIssue
    {
        [JsonPropertyName("custom_fields")]
        public List<WeldingCustomField> CustomFields { get; set; }

        [JsonPropertyName("attachments")]
        public List<WeldingAttachment> Attachments { get; set; }
    }

    public class WeldingCustomField
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("value")]
        public JsonElement Value { get; set; }
    }

    public class WeldingAttachment
    {
        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("content_url")]
        public string ContentUrl { get; set; }
    }



    public class RequestIssueResponse
    {
        [JsonPropertyName("issue")]
        public RequestIssue Issue { get; set; }
    }

    public class RequestIssue
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        //[JsonPropertyName("subject")]
        //public string Subject { get; set; }

        [JsonPropertyName("custom_fields")]
        public List<RequestCustomField> CustomFields { get; set; }

        [JsonPropertyName("tracker")]
        public RequestTracker Tracker { get; set; }

        [JsonPropertyName("start_date")]
        public string StartDate { get; set; }
    }

    public class RequestCustomField
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("value")]
        public JsonElement Value { get; set; }
    }

    public class RequestTracker
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }
    }

    public class RequestReportData
    {
        public string Name { get; set; }
        public string CustomerName { get; set; }
        //public string CustomerPosition { get; set; }
        //public string CustomerPhone { get; set; }
        public string CustomerEmail { get; set; }
        public string Theme { get; set; }
        public string Aim { get; set; }
        public string RequestDate { get; set; }
        public string CuratorName { get; set; }
        //public string CuratorPosition { get; set; }
        //public string CuratorPhone { get; set; }
        public string CuratorEmail { get; set; }
        public string PlanStartDateText { get; set; }
        public string PlanEndDateText { get; set; }
        public string Cost { get; set; }
        public string CostText { get; set; }
        public string OwnCost { get; set; }
        public string SubCost { get; set; }
        public string MaterialCost { get; set; }
        public string OtherCost { get; set; }
    }
}
