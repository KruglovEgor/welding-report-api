using System.Collections.Concurrent;
using Microsoft.Extensions.Options;
using welding_report.Models;
using welding_report.Services.Welding;

namespace welding_report.Services
{
    public interface IEmailServiceFactory
    {
        IWeldingEmailService CreateWeldingEmailService(string apiKey);
    }

    public class EmailServiceFactory : IEmailServiceFactory
    {
        private readonly IOptions<EmailSettings> _emailSettings;
        private readonly ILoggerFactory _loggerFactory;
        private readonly IRedmineServiceFactory _redmineServiceFactory;

        private readonly ConcurrentDictionary<string, IWeldingEmailService> _emailServiceCache =
            new ConcurrentDictionary<string, IWeldingEmailService>();

        public EmailServiceFactory(
            IOptions<EmailSettings> emailSettings,
            ILoggerFactory loggerFactory,
            IRedmineServiceFactory redmineServiceFactory)
        {
            _emailSettings = emailSettings;
            _loggerFactory = loggerFactory;
            _redmineServiceFactory = redmineServiceFactory;
        }

        public IWeldingEmailService CreateWeldingEmailService(string apiKey)
        {
            return _emailServiceCache.GetOrAdd(apiKey, key => {
                // Используем фабрику RedmineService для создания соответствующего сервиса
                var redmineService = _redmineServiceFactory.CreateWeldingService(key);
                return new WeldingEmailService(_emailSettings, redmineService);
            });
        }
    }
}
