using System.Text.Json.Serialization;

namespace welding_report.Models
{
    public class RedmineSettings
    {
        public string WeldingUrl { get; set; }
        public string WeldingApiKey { get; set; }
        public string RequestUrl { get; set; }
        public string RequestApiKey { get; set; }
        public string SuprUrl { get; set; }
        public string SuprApiKey { get; set; }

    }

    public class AccountInfo
    {
        [JsonPropertyName("user")]
        public UserInfo User { get; set; }
    }

    public class UserInfo
    {
        [JsonPropertyName("firstname")]
        public string FirstName { get; set; }
        [JsonPropertyName("lastname")]
        public string LastName { get; set; }

        [JsonPropertyName("mail")]
        public string Mail { get; set; }
    }

}
