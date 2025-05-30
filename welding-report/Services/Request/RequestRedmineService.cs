using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using System.Text.Json;
using welding_report.Models;
using System.Globalization;
using DocumentFormat.OpenXml;


namespace welding_report.Services.Request
{
    public interface IRequestRedmineService
    {
        Task<RequestReportData> GetRequestReportDataAsync(int issueId);
    }

    public class RequestRedmineService : IRequestRedmineService
    {
        private readonly HttpClient _httpClient;
        private readonly RedmineSettings _settings;
        private readonly ILogger<RedmineService> _logger;


        public RequestRedmineService(
            IHttpClientFactory httpClientFactory,
            IOptions<RedmineSettings> redmineSettings,
            ILogger<RedmineService> logger,
            string apiKey)
        {
            _settings = redmineSettings.Value;
            _logger = logger;

            // Получаем клиент из фабрики вместо создания нового
            _httpClient = httpClientFactory.CreateClient("Request"); // Используем именованный клиент

            // Настройка только для этого экземпляра
            _httpClient.BaseAddress = new Uri(_settings.RequestUrl);
            _httpClient.DefaultRequestHeaders.Accept.Clear(); // Очищаем перед добавлением
            _httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", apiKey);
        }


        public async Task<RequestReportData> GetRequestReportDataAsync(int issueId)
        {

            var reportData = new RequestReportData();

            var issueResponse = await _httpClient.GetFromJsonAsync<RequestIssueResponse>($"issues/{issueId}.json");
            if (issueResponse?.Issue == null)
            {
                throw new Exception("Акт не найден");
            }

            reportData.Name = $"{issueResponse.Issue.Tracker.Name}-{issueId}";

            var startDate = issueResponse.Issue.StartDate;
            if (DateTime.TryParseExact(startDate,
                             "yyyy-MM-dd",
                             CultureInfo.InvariantCulture,
                             DateTimeStyles.None,
                             out DateTime parsedStartDate))
            {
                reportData.RequestDate = parsedStartDate.ToString("dd.MM.yyyy");
            }
            else
            {
                reportData.RequestDate = startDate;
            }

            reportData.Theme = issueResponse.Issue.Subject;

            foreach (var field in issueResponse.Issue.CustomFields)
            {
                if (field.Name == "Куратор от НПЗ")
                {
                    var stringValue = field.Value.GetString();
                    if (int.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                    {
                        var userInfo = await GetUserInfoAsync(parsedValue);
                        if (userInfo?.User != null)
                        {
                            reportData.CustomerName = $"{userInfo.User.LastName} {userInfo.User.FirstName}";
                            reportData.CustomerEmail = userInfo.User.Mail;
                        }
                    }
                }

                if (field.Name == "Куратор заявки ЛИНК")
                {
                    var stringValue = field.Value.GetString();
                    if (int.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                    {
                        var userInfo = await GetUserInfoAsync(parsedValue);
                        if (userInfo?.User != null)
                        {
                            reportData.CuratorName = $"{userInfo.User.LastName} {userInfo.User.FirstName}";
                            reportData.CuratorEmail = userInfo.User.Mail;
                        }
                    }
                }

                if (field.Name == "Цель работы")
                {
                    var stringValue = field.Value.GetString();
                    reportData.Aim = stringValue;
                }

                if (field.Name == "Ожидаемая дата начала")
                {
                    var stringValue = field.Value.GetString();
                    if (DateTime.TryParseExact(stringValue,
                             "yyyy-MM-dd",
                             CultureInfo.InvariantCulture,
                             DateTimeStyles.None,
                             out DateTime parsedDate))
                    {
                        reportData.PlanStartDateText = parsedDate.ToString("dd.MM.yyyy");
                    }
                    else
                    {
                        // Обработка некорректного формата, можно оставить оригинальную строку или задать ошибку
                        reportData.PlanStartDateText = stringValue;
                    }

                }

                if (field.Name == "Ожидаемый срок завершения")
                {
                    var stringValue = field.Value.GetString();
                    if (DateTime.TryParseExact(stringValue,
                             "yyyy-MM-dd",
                             CultureInfo.InvariantCulture,
                             DateTimeStyles.None,
                             out DateTime parsedDate))
                    {
                        reportData.PlanEndDateText = parsedDate.ToString("dd.MM.yyyy");
                    }
                    else
                    {
                        reportData.PlanEndDateText = stringValue;
                    }
                }

                if (field.Name == "Сумма, руб")
                {
                    var stringValue = field.Value.GetString();
                    reportData.Cost = stringValue;
                }

                if (field.Name == "Объем работ силами ЛИНК, руб")
                {
                    var stringValue = field.Value.GetString();
                    reportData.OwnCost = stringValue;
                }

                if (field.Name == "Объем субподряда, руб")
                {
                    var stringValue = field.Value.GetString();
                    reportData.SubCost = stringValue;
                }

                if (field.Name == "Материальные затраты, руб")
                {
                    var stringValue = field.Value.GetString();
                    reportData.MaterialCost = stringValue;
                }
                if (field.Name == "Прочие затраты, руб")
                {
                    var stringValue = field.Value.GetString();
                    reportData.OtherCost = stringValue;
                }
            }

            return reportData;
        }

        private async Task<AccountInfo> GetUserInfoAsync(int userId)
        {
            var response = await _httpClient.GetAsync($"users/{userId}.json");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<AccountInfo>();
        }

    }
}
