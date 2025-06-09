using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using System.Text.Json;
using welding_report.Models;
using System.Globalization;
using DocumentFormat.OpenXml;


namespace welding_report.Services.Supr
{
    public interface ISuprRedmineService
    {
        Task<SuprGroupReportData> GetSuprGroupReportDataAsync(string projectIdentifier, int applicationNumber);
    }

    public class SuprRedmineService : ISuprRedmineService
    {
        private readonly HttpClient _httpClient;
        private readonly RedmineSettings _settings;
        private readonly ILogger<SuprRedmineService> _logger;

        public SuprRedmineService(
            IHttpClientFactory httpClientFactory,
            IOptions<RedmineSettings> redmineSettings,
            ILogger<SuprRedmineService> logger,
            string apiKey)
        {

            _settings = redmineSettings.Value;
            _logger = logger;

            // Получаем клиент из фабрики вместо создания нового
            _httpClient = httpClientFactory.CreateClient("Supr"); // Используем именованный клиент

            // Настройка только для этого экземпляра
            _httpClient.BaseAddress = new Uri(_settings.SuprUrl);
            _httpClient.DefaultRequestHeaders.Accept.Clear(); // Очищаем перед добавлением
            _httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);
        }


        public async Task<SuprGroupReportData> GetSuprGroupReportDataAsync(string projectIdentifier, int applicationNumber)
        {

            var reportData = new SuprGroupReportData();

            reportData.ApplicationNumber = applicationNumber;

            var projectResponse = await _httpClient.GetAsync($"projects/{projectIdentifier}.json");
            projectResponse.EnsureSuccessStatusCode();

            var projectResponseData = await projectResponse.Content.ReadFromJsonAsync<SuprProjectRespose>();
            if (projectResponseData?.Project != null)
            {
                string projectDescription = projectResponseData.Project.Description;
                int index = projectDescription.IndexOf("от ");
                if (index != -1)
                {
                    reportData.CustomerCompany = projectDescription.Substring(index+3).Trim();
                }

                foreach(var field in projectResponseData.Project.CustomFields)
                {
                    switch (field.Id)
                    {
                        case 61:
                            reportData.ContractNumber = field.Value.GetString().Trim();
                            break;
                        case 62:
                            string result = "";
                            var fullName = field.Value.GetString();
                            var parts = fullName.Split(' ');
                            if (parts.Length >= 2)
                            {
                                for (int i = 1; i < parts.Length; i++)
                                {
                                    if (!string.IsNullOrWhiteSpace(parts[i]))
                                        result += parts[i][0] + ".";
                                }
                            }
                            result += " " + parts[0];
                            reportData.CustomerRepresentative = result.Trim();
                            break;
                    }
                }

            }

            var response = await _httpClient.GetAsync($"projects/{projectIdentifier}/issues.json?cf_38={applicationNumber}");
            response.EnsureSuccessStatusCode();

            var issuesResponse = await response.Content.ReadFromJsonAsync<SuprIssueListResponse>();
            if (issuesResponse?.Issues == null || issuesResponse.Issues.Count == 0)
                return reportData;

            bool first = true;
            int issuesCount = issuesResponse.Issues.Count;

            reportData.suprIssueReportDatas = new Dictionary<int, SuprIssueReportData>();


            foreach (SuprIssue issue in issuesResponse.Issues)
            {
                if (first)
                {
                    reportData.Factory = issue.Project.Name.Split(' ')[0];

                    // Парсинг даты создания
                    if (DateTime.TryParse(issue.CreateDate, out DateTime createdDate))
                    {
                        reportData.CreateDate = createdDate;
                    }
                    else
                    {
                        reportData.CreateDate = DateTime.Now;
                    }

                    first = false;
                }

                SuprIssueReportData suprIssueReportData = new SuprIssueReportData();
                int number = 0;

                suprIssueReportData.Detail = issue.Subject;
                suprIssueReportData.Priority = issue.Priority.Name;

                string mark = "";
                string manufacturer = "";

                foreach (var field in issue.CustomFields)
                {
                    switch (field.Id)
                    {
                        case 58: // Номер п/п
                            if (int.TryParse(field.Value.GetString(), out int parsedNumber))
                            {
                                number = parsedNumber;
                                if (!reportData.suprIssueReportDatas.ContainsKey(number))
                                {
                                    // переходим к следующему кастомному полю
                                    continue;
                                }
                            }

                            for (int i = issuesCount; i >= 1; i--)
                            {
                                if (!reportData.suprIssueReportDatas.ContainsKey(i))
                                {
                                    number = i;
                                    break;
                                }
                            }

                            break;

                        case 20: // Период сканирования
                            suprIssueReportData.ScanningPeriod = field.Value.GetString();
                            break;
                        case 54: // Состояние
                            suprIssueReportData.Condition = field.Value.GetString();
                            break;
                        case 55: // Вид работ
                            var values = new List<string>();
                            foreach (var value in field.Value.EnumerateArray())
                            {
                                values.Add(value.GetString());
                            }
                            suprIssueReportData.JobType = string.Join(", ", values);
                            break;
                        case 1: // Название установки
                            suprIssueReportData.InstallationName = field.Value.GetString();
                            break;
                        case 18: // Тех. позиция
                            suprIssueReportData.TechPositionName = field.Value.GetString();
                            break;
                        case 53: // Номер оборудования
                            suprIssueReportData.EquipmentUnitNumber = field.Value.GetString();
                            break;
                        case 3: // Марка
                            mark = field.Value.GetString();
                            break;
                        case 40: // Изготовитель
                            manufacturer = field.Value.GetString();
                            break;
                    }
                }

                if (string.IsNullOrEmpty(mark))
                {
                    suprIssueReportData.MarkAndManufacturer = manufacturer;
                }
                else
                {
                    suprIssueReportData.MarkAndManufacturer = string.IsNullOrEmpty(manufacturer) ? mark : $"{mark}, {manufacturer}";
                }

                // Проверяем, является ли текущая дата создания более ранней
                if (DateTime.TryParse(issue.CreateDate, out DateTime currentIssueDate) &&
                    currentIssueDate < reportData.CreateDate)
                {
                    reportData.CreateDate = currentIssueDate;
                }

                reportData.suprIssueReportDatas[number] = suprIssueReportData;
            }

            return reportData;
        }
    }
}
