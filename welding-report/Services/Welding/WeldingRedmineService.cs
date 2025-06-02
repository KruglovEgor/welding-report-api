using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using System.Text.Json;
using welding_report.Models;
using System.Globalization;
using DocumentFormat.OpenXml;



namespace welding_report.Services.Welding
{
    public interface IWeldingRedmineService
    {
        Task<WeldingIssueReportData> GetWeldingIssueDataAsync(int projectIdentifier, int parentIssueId);
        Task<WeldingProjectReportData> GetProjectReportDataAsync(int projectIdentifier);
        Task<AccountInfo> GetCurrentUserInfoAsync();

    }

    public class WeldingRedmineService : IWeldingRedmineService
    {
        private readonly HttpClient _httpClient;
        private readonly RedmineSettings _settings;
        private readonly ILogger<WeldingRedmineService> _logger;

        public WeldingRedmineService(
            IHttpClientFactory httpClientFactory,
            IOptions<RedmineSettings> redmineSettings,
            ILogger<WeldingRedmineService> logger,
            string apiKey)
        {
            _settings = redmineSettings.Value;
            _logger = logger;

            // Получаем клиент из фабрики вместо создания нового
            _httpClient = httpClientFactory.CreateClient("Welding"); // Используем именованный клиент

            // Настройка только для этого экземпляра
            _httpClient.BaseAddress = new Uri(_settings.WeldingUrl);
            _httpClient.DefaultRequestHeaders.Accept.Clear(); // Очищаем перед добавлением
            _httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);
        }

      
        public async Task<WeldingIssueReportData> GetWeldingIssueDataAsync(int projectIdentifier, int parentIssueId)
        {
            var issueReportData = new WeldingIssueReportData();

            // Получение данных родительского акта
            var parentResponse = await _httpClient.GetFromJsonAsync<WeldingIssueResponse>($"issues/{parentIssueId}.json");
            if (parentResponse?.Issue == null)
                throw new Exception("Родительский акт не найден");

            issueReportData.ReportNumber = parentResponse.Issue.Subject;

            foreach (var field in parentResponse.Issue.CustomFields)
            {
                if (field.Name == "Количество стыков")
                {
                    var stringValue = field.Value.GetString();
                    if (int.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                    {
                        issueReportData.JointsCountPlan = parsedValue;
                        _logger.LogInformation($"Parsed JointsCountPlan: {issueReportData.JointsCountFact}");
                    }
                }

                if (field.Name == "Дюймы_ПЛАН")
                {
                    var stringValue = field.Value.GetString();
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                    {
                        issueReportData.DiametrInchesPlan = Math.Round(parsedValue, 2);
                        _logger.LogInformation($"Parsed DiametrInchesPlan: {issueReportData.JointsCountFact}");
                    }
                }

                if (field.Name == "Дюймы_ФАКТ")
                {
                    var stringValue = field.Value.GetString();
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                    {
                        issueReportData.DiametrInchesFact = Math.Round(parsedValue, 2);
                        _logger.LogInformation($"Parsed DiametrInchesFact: {issueReportData.JointsCountFact}");
                    }
                }
            }

            // Получение дочерних задач
            var childrenResponse = await _httpClient.GetFromJsonAsync<RedmineChildIssueListResponse>($"projects/{projectIdentifier}/issues.json?parent_id={parentIssueId}&status_id=*&include=attachments");
            if (childrenResponse?.Issues == null)
                return issueReportData;


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
                            group.DiameterInches = Math.Round(parsedValue, 2);
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

                    if (field.Name == "Количество стыков")
                    {
                        var stringValue = field.Value.GetString();
                        if (int.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            group.JointsCount = parsedValue;
                            issueReportData.JointsCountFact += parsedValue;
                            _logger.LogInformation($"Parsed JointsCount: {group.DiameterInches}");
                        }
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

                        // Проверка существующей записи
                        if (entry.JointPhotoMap.TryGetValue(joints, out var photoList))
                        {
                            photoList.Add(attachment.ContentUrl);
                        }
                        else
                        {
                            entry.JointPhotoMap[joints] = new List<string> { attachment.ContentUrl };
                        }
                    }
                    catch { /* Игнорируем битые вложения */ }
                }

                group.Entries = entries.Values.ToList();
                _logger.LogInformation($"Entries: {entries.Values.ToList()}");
                issueReportData.Groups.Add(group);
            }

            return issueReportData;
        }


        public async Task<WeldingProjectReportData> GetProjectReportDataAsync(int projectIdentifier)
        {
            var projectReport = new WeldingProjectReportData { Identifier = projectIdentifier };

            var projectResponse = await _httpClient.GetAsync(
                $"projects/{projectIdentifier}.json"
            );



            var parsedProjectResponse = await projectResponse.Content.ReadFromJsonAsync<WeldingProjectResponse>();
            if (parsedProjectResponse?.Project?.Name != null)
            {
                projectReport.Name = parsedProjectResponse.Project.Name;
            }
            else
            {
                projectReport.Name = projectIdentifier.ToString();
            }

            // Получаем все акты проекта (трекер ID=1)
            var response = await _httpClient.GetAsync(
                $"projects/{projectIdentifier}/issues.json?tracker_id=1&status_id=*"
            );
            response.EnsureSuccessStatusCode();

            var issuesResponse = await response.Content.ReadFromJsonAsync<WeldingIssueListResponse>();
            if (issuesResponse?.Issues == null) return projectReport;

            // Для каждого акта собираем данные
            foreach (var actIssue in issuesResponse.Issues)
            {
                var actData = await GetWeldingIssueDataAsync(projectIdentifier, actIssue.Id);
                projectReport.Acts.Add(actData);
            }

            return projectReport;
        }
        public async Task<AccountInfo> GetCurrentUserInfoAsync()
        {
            var response = await _httpClient.GetAsync("/my/account.json");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<AccountInfo>();
        }
    }
}
