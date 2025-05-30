using System.Collections.Concurrent;
using Microsoft.Extensions.Options;
using welding_report.Models;
using welding_report.Services.Request;

namespace welding_report.Services
{
    public interface IRedmineServiceFactory
    {
        IRequestRedmineService CreateRequestService(string apiKey);
        // Другие сервисы...
    }

    public class RedmineServiceFactory : IRedmineServiceFactory
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IOptions<RedmineSettings> _settings;
        private readonly ILoggerFactory _loggerFactory;

        private ConcurrentDictionary<string, IRequestRedmineService> _serviceCache =
        new ConcurrentDictionary<string, IRequestRedmineService>();

        public RedmineServiceFactory(
            IHttpClientFactory httpClientFactory,
            IOptions<RedmineSettings> settings,
            ILoggerFactory loggerFactory)
        {
            _httpClientFactory = httpClientFactory;
            _settings = settings;
            _loggerFactory = loggerFactory;
        }

        public IRequestRedmineService CreateRequestService(string apiKey)
        {
            // Возвращаем существующий сервис из кэша или создаем новый
            return _serviceCache.GetOrAdd(apiKey, key => {
                var logger = _loggerFactory.CreateLogger<RedmineService>();
                return new RequestRedmineService(
                    _httpClientFactory.CreateClient(),
                    _settings,
                    logger,
                    key);
            });
        }



    }
}
