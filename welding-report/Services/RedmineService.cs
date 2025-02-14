using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using System.Text.Json;
using welding_report.Models;

namespace welding_report.Services
{
    public interface IRedmineService
    {
        Task<T> GetIssueAsync<T>(int issueId);
        Task<T> GetChildIssuesAsync<T>(int parentId);
    }

    public class RedmineService : IRedmineService
    {
        private readonly HttpClient _httpClient;
        private readonly RedmineSettings _settings;

        public RedmineService(
            HttpClient httpClient,
            IOptions<RedmineSettings> redmineSettings)
        {
            _httpClient = httpClient;
            _settings = redmineSettings.Value;

            // Настройка базового URL и заголовков
            _httpClient.BaseAddress = new Uri(_settings.BaseUrl);
            _httpClient.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", _settings.ApiKey);
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


    }
}
