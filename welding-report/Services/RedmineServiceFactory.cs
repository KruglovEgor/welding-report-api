using System.Collections.Concurrent;
using Microsoft.Extensions.Options;
using welding_report.Models;
using welding_report.Services.Request;
using welding_report.Services.Supr;
using welding_report.Services.Welding;

namespace welding_report.Services
{
    public interface IRedmineServiceFactory
    {
        IRequestRedmineService CreateRequestService(string apiKey);
        IWeldingRedmineService CreateWeldingService(string apiKey);
        ISuprRedmineService CreateSuprService(string apiKey);

        IWeldingExcelReportGenerator CreateWeldingExcelGenerator(string apiKey);
    }

    public class RedmineServiceFactory : IRedmineServiceFactory
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IOptions<RedmineSettings> _settings;
        private readonly ILoggerFactory _loggerFactory;
        private readonly IOptions<AppSettings> _appSettings;


        private readonly ConcurrentDictionary<string, IRequestRedmineService> _requestServiceCache =
             new ConcurrentDictionary<string, IRequestRedmineService>();

        private readonly ConcurrentDictionary<string, IWeldingRedmineService> _weldingServiceCache =
            new ConcurrentDictionary<string, IWeldingRedmineService>();

        private readonly ConcurrentDictionary<string, ISuprRedmineService> _suprServiceCache =
            new ConcurrentDictionary<string, ISuprRedmineService>();

        private readonly ConcurrentDictionary<string, IWeldingExcelReportGenerator> _excelGeneratorCache =
            new ConcurrentDictionary<string, IWeldingExcelReportGenerator>();

        public RedmineServiceFactory(
            IHttpClientFactory httpClientFactory,
            IOptions<RedmineSettings> settings,
            ILoggerFactory loggerFactory,
            IOptions<AppSettings> appSettings
            )
        {
            _httpClientFactory = httpClientFactory;
            _settings = settings;
            _loggerFactory = loggerFactory;
            _appSettings = appSettings;
        }

        public IRequestRedmineService CreateRequestService(string apiKey)
        {
            return _requestServiceCache.GetOrAdd(apiKey, key => {
                var logger = _loggerFactory.CreateLogger<RequestRedmineService>();
                return new RequestRedmineService(
                    _httpClientFactory,
                    _settings,
                    logger,
                    key);
            });
        }

        public IWeldingRedmineService CreateWeldingService(string apiKey)
        {
            return _weldingServiceCache.GetOrAdd(apiKey, key => {
                var logger = _loggerFactory.CreateLogger<WeldingRedmineService>();
                return new WeldingRedmineService(
                    _httpClientFactory,
                    _settings,
                    logger,
                    key);
            });
        }

        public ISuprRedmineService CreateSuprService(string apiKey)
        {
            return _suprServiceCache.GetOrAdd(apiKey, key => {
                var logger = _loggerFactory.CreateLogger<SuprRedmineService>();
                return new SuprRedmineService(
                    _httpClientFactory,
                    _settings,
                    logger,
                    key);
            });
        }

        public IWeldingExcelReportGenerator CreateWeldingExcelGenerator(string apiKey)
        {
            return _excelGeneratorCache.GetOrAdd(apiKey, key => {
                var logger = _loggerFactory.CreateLogger<WeldingExcelReportGenerator>();
                return new WeldingExcelReportGenerator(
                    logger,
                    _appSettings,
                    key);
            });
        }

    }
}
