using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using System.Text.Json;
using welding_report.Models;
using System.Linq;
using System.Globalization;


namespace welding_report.Services
{
    public interface IRedmineService
    {
        Task<T> GetIssueAsync<T>(int issueId);
        Task<T> GetChildIssuesAsync<T>(int parentId);
        Task<RedmineReportData> GetReportDataAsync(int parentIssueId);
    }

    public class RedmineService : IRedmineService
    {
        private readonly HttpClient _httpClient;
        private readonly RedmineSettings _settings;
        private readonly ILogger<RedmineService> _logger;

        public RedmineService(
            HttpClient httpClient,
            IOptions<RedmineSettings> redmineSettings,
            ILogger<RedmineService> logger)
        {
            _httpClient = httpClient;
            _settings = redmineSettings.Value;
            _logger = logger;

            // Настройка базового URL и заголовков
            _httpClient.BaseAddress = new Uri(_settings.BaseUrl);
            _httpClient.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", _settings.ApiKey);
            _logger = logger;
        }

        public async Task<T> GetIssueAsync<T>(int issueId)
        {
            var response = await _httpClient.GetAsync($"issues/{issueId}.json");
            response.EnsureSuccessStatusCode();
            //return await response.Content.ReadFromJsonAsync<T>();
            return JsonSerializer.Deserialize<dynamic>(await response.Content.ReadAsStringAsync());
        }

        public async Task<T> GetChildIssuesAsync<T>(int parentId)
        {
            var response = await _httpClient.GetAsync($"issues.json?parent_id={parentId}&status_id=*&include=attachments");
            response.EnsureSuccessStatusCode();
            //return await response.Content.ReadFromJsonAsync<T>();
            return JsonSerializer.Deserialize<dynamic>(await response.Content.ReadAsStringAsync());
        }

        public async Task<RedmineReportData> GetReportDataAsync(int parentIssueId)
        {
            var reportData = new RedmineReportData();

            // Получение данных родительского акта
            var parentResponse = await _httpClient.GetFromJsonAsync<RedmineIssueResponse>($"issues/{parentIssueId}.json");
            if (parentResponse?.Issue == null)
                throw new Exception("Родительский акт не найден");

            reportData.ReportNumber = parentResponse.Issue.Subject;

            // Получение дочерних задач
            var childrenResponse = await _httpClient.GetFromJsonAsync<RedmineIssueListResponse>($"issues.json?parent_id={parentIssueId}&status_id=*&include=attachments");
            if (childrenResponse?.Issues == null)
                return reportData;
            

            foreach (var child in childrenResponse.Issues)
            {
                var group = new JointGroup();

                // Парсинг данных для стыка
                foreach (var field in child.CustomFields)
                {
                    if (field.Name == "Наружный диаметр")
                    {
                        var stringValue = field.Value.GetString();
                        if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            group.DiameterMm = parsedValue;
                            _logger.LogInformation($"Parsed Diam: {group.DiameterMm}");
                        }
                    }
                        

                    if (field.Name == "Дюймы_ФАКТ")
                    {
                        var stringValue = field.Value.GetString();
                        if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            group.DiameterInches = parsedValue;
                            _logger.LogInformation($"Parsed Inches: {group.DiameterInches}");
                        }
                    }

                    if (field.Name == "Пункт акта")
                    {
                        var stringValue = field.Value.GetString();

                        if (int.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            group.ActParagraph = parsedValue;
                            _logger.LogInformation($"Parsed ActParagraph: {group.ActParagraph}");
                        }
                    }

                    if (field.Name == "Тип оборудования")
                    {
                        group.EquipmentType = field.Value.GetString();
                        _logger.LogInformation($"Parsed EquipmentType: {group.EquipmentType}");
                    }

                    if (field.Name == "№ трубопровода/аппарата")
                    {
                        group.PipelineNumber = field.Value.GetString();
                        _logger.LogInformation($"Parsed PipelineNumber: {group.PipelineNumber}");
                    }


                }

                // Обработка вложений
                var entries = new Dictionary<string, JointEntry>();
                foreach (var attachment in child.Attachments)
                {
                    try
                    {
                        using var descDoc = JsonDocument.Parse(attachment.Description);
                        var root = descDoc.RootElement;

                        var contractor = root.GetProperty("Contractor").GetString();
                        var joints = root.GetProperty("MarkedJoints").GetString();

                        if (!entries.TryGetValue(contractor, out var entry))
                        {
                            entry = new JointEntry { Contractor = contractor };
                            entries[contractor] = entry;
                        }

                        entry.JointNumbers = string.IsNullOrEmpty(entry.JointNumbers)
                            ? joints
                            : $"{entry.JointNumbers}, {joints}";

                        entry.PhotoUrls.Add(attachment.ContentUrl);
                    }
                    catch { /* Игнорируем битые вложения */ }
                }

                group.Entries = entries.Values.ToList();
                reportData.Groups.Add(group);
            }

            return reportData;
        }

    }
}
