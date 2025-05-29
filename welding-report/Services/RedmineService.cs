using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using System.Text.Json;
using welding_report.Models;
using System.Globalization;
using DocumentFormat.OpenXml;


namespace welding_report.Services
{
    public interface IRedmineService
    {
        Task<WeldingIssueReportData> GetWeldingIssueDataAsync(int projectIdentifier, int parentIssueId);
        Task<AccountInfo> GetCurrentUserInfoAsync();
        Task<WeldingProjectReportData> GetProjectReportDataAsync(int projectIdentifier);
        Task<RequestReportData> GetRequestReportDataAsync(int issueId);

        Task<SuprGroupReportData> GetSuprGroupReportDataAsync(string projectIdentifier, int applicationNumber);
        void SetApiKey(string apiKey);
        void SetContext(string context);
        void SetHttpClient();
    }

    public class RedmineService : IRedmineService
    {
        private readonly HttpClient _httpClient;
        private readonly RedmineSettings _settings;
        private readonly ILogger<RedmineService> _logger;

        private string _apiKey = "";
        private string _context = "";

        public RedmineService(
            HttpClient httpClient,
            IOptions<RedmineSettings> redmineSettings,
            ILogger<RedmineService> logger)
        {
            //_httpClient = httpClient;

            _settings = redmineSettings.Value;
            _logger = logger;

            HttpClientHandler clientHandler = new HttpClientHandler();
            clientHandler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => { return true; };
            _httpClient = new HttpClient(clientHandler);
        }

        public void SetApiKey(string apiKey)
        {
            _apiKey = apiKey;
        }

        public void SetContext(string context)
        {
            _context = context;
        }

        public void SetHttpClient()
        {
            if (_context == "welding")
            {
                _httpClient.BaseAddress = new Uri(_settings.WeldingUrl);
            }
            else if (_context == "request")
            {
                _httpClient.BaseAddress = new Uri(_settings.RequestUrl);
            }
            else if (_context == "supr")
            {
                _httpClient.BaseAddress = new Uri(_settings.SuprUrl);
            }
            
                _httpClient.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", _apiKey);
        }

        public async Task<RequestReportData> GetRequestReportDataAsync(int issueId)
        {
            if (_context != "request")
            {
                _context = "request";
                SetHttpClient();
            }

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


        public async Task<WeldingIssueReportData> GetWeldingIssueDataAsync(int projectIdentifier, int parentIssueId)
        {
            if (_context != "welding")
            {
                _context = "welding";
                SetHttpClient();
            }
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
            if (_context != "welding")
            {
                _context = "welding";
                SetHttpClient();
            }

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

        private async Task<AccountInfo> GetUserInfoAsync(int userId)
        {
            var response = await _httpClient.GetAsync($"users/{userId}.json");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<AccountInfo>();
        }


        public async Task<SuprGroupReportData> GetSuprGroupReportDataAsync(string projectIdentifier, int applicationNumber)
        {
            if (_context != "supr")
            {
                _context = "supr";
                SetHttpClient();
            }

            var reportData = new SuprGroupReportData();

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

                    string mark = "";
                    string manufacturer = "";

                    foreach (var field in issue.CustomFields)
                    {
                        switch (field.Id)
                        {
                            case 1: // Название установки
                                reportData.InstallationName = field.Value.GetString();
                                break;
                            case 18: // Тех. позиция
                                reportData.TechPositionName = field.Value.GetString();
                                break;
                            case 53: // Номер оборудования
                                reportData.EquipmentUnitNumber = field.Value.GetString();
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
                        reportData.MarkAndManufacturer = manufacturer;
                    }
                    else
                    {
                        reportData.MarkAndManufacturer = string.IsNullOrEmpty(manufacturer) ? mark : $"{mark}, {manufacturer}";
                    }

                    first = false;
                }

                SuprIssueReportData suprIssueReportData = new SuprIssueReportData();
                int number = 0;

                suprIssueReportData.Detail = issue.Subject;
                suprIssueReportData.Priority = issue.Priority.Name;

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
                    }
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
