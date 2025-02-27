using System.Text.Json.Serialization;

namespace welding_report.Models
{
    public class RedmineSettings
    {
        public string BaseUrl { get; set; }
        public string ApiKey { get; set; }
    }

    public class RedmineAccountInfo
    {
        [JsonPropertyName("user")]
        public RedmineUserInfo User { get; set; }
    }

    public class RedmineUserInfo
    {
        [JsonPropertyName("firstname")]
        public string FirstName { get; set; }
        [JsonPropertyName("lastname")]
        public string LastName { get; set; }

        [JsonPropertyName("mail")]
        public string Mail { get; set; }
    }
}
