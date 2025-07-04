using System.Text.Json.Serialization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Security.Cryptography.Pkcs;

namespace welding_report.Models
{
    public class WeldingProjectReportData
    {
        public string Name { get; set; }
        public int Identifier { get; set; }
        public List<WeldingIssueReportData> Acts { get; set; } = new();
    }

    public class WeldingProjectResponse
    {
        [JsonPropertyName("project")]
        public Project Project { get; set; }
    }

    public class Project
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

    }

    public class WeldingIssueReportData
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
        public List<CustomField> CustomFields { get; set; }
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
        public List<CustomField> CustomFields { get; set; }

        [JsonPropertyName("attachments")]
        public List<WeldingAttachment> Attachments { get; set; }
    }

    public class CustomField
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

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
        public List<CustomField> CustomFields { get; set; }

        [JsonPropertyName("tracker")]
        public RequestTracker Tracker { get; set; }

        [JsonPropertyName("start_date")]
        public string StartDate { get; set; }
        [JsonPropertyName("subject")]
        public string Subject { get; set; }
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


    public class SuprProjectRespose
    {
        [JsonPropertyName("project")]
        public SuprProject Project { get; set; }

    }

    public class SuprProject
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [JsonPropertyName("description")]
        public string Description { get; set; }
        [JsonPropertyName("custom_fields")]
        public List<CustomField> CustomFields { get; set; }

    }
    
    public class SuprIssueListResponse
    {
        [JsonPropertyName("issues")]
        public List<SuprIssue> Issues { get; set; }

    }


    public class SuprIssue
    {
        [JsonPropertyName("project")]
        public Project Project { get; set; }

        [JsonPropertyName("subject")]
        public string Subject { get; set; }

        [JsonPropertyName("custom_fields")]
        public List<CustomField> CustomFields { get; set; }

        [JsonPropertyName("priority")]
        public Priority Priority { get; set; }


        [JsonPropertyName("created_on")]
        public string CreateDate { get; set; }
    }


    public class Priority
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }
    }

    public class SuprGroupReportData
    {
        public Dictionary<int, SuprIssueReportData> suprIssueReportDatas { get; set; }

        public string Factory { get; set; }

        //public string EquipmentType { get; set; }

        public DateTime CreateDate { get; set; }

        public int ApplicationNumber { get; set; }
        public string CustomerCompany { get; set; }
        public string CustomerRepresentative { get; set; }
        public string ContractNumber { get; set; }
    }


    public class SuprIssueReportData
    {
        public string Detail { get; set; }
        public string ScanningPeriod { get; set; }
        public string Condition { get; set; }
        public string Priority { get; set; }
        public string JobType { get; set; }
        public string EquipmentType { get; set; }
        public string InstallationName { get; set; }

        public string TechPositionName { get; set; }

        public string EquipmentUnitNumber { get; set; }

        public string MarkAndManufacturer { get; set; }
    }
}