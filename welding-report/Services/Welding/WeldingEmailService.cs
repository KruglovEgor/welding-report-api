using Microsoft.Extensions.Options;
using System.Net.Mail;
using System.Net;
using welding_report.Models;
using welding_report.Services.Welding;

namespace welding_report.Services.Welding
{
    public interface IWeldingEmailService
    {
        Task SendReportAsync(string recipientEmail, string subject, string body, byte[] attachment, string attachmentName);
        Task SendRedmineReportAsync(byte[] reportBytes, string name, string context);
    }

    public class WeldingEmailService : IWeldingEmailService
    {
        private readonly EmailSettings _emailSettings;
        private readonly IWeldingRedmineService _weldingRedmineService;

        public WeldingEmailService(
            IOptions<EmailSettings> emailSettings,
            IWeldingRedmineService weldingRedmineService)
        {
            _emailSettings = emailSettings.Value;
            _weldingRedmineService = weldingRedmineService;
        }

        public async Task SendReportAsync(string recipientEmail, string subject, string body, byte[] attachment, string attachmentName)
        {
            using var smtpClient = new SmtpClient(_emailSettings.SmtpServer)
            {
                Port = _emailSettings.Port,
                Credentials = new NetworkCredential(_emailSettings.Username, _emailSettings.Password),
                EnableSsl = true,
                Timeout = 15000 // 15 секунд, чтобы не зависало
            };

            using var message = new MailMessage
            {
                From = new MailAddress(_emailSettings.SenderEmail, _emailSettings.SenderName),
                Subject = subject,
                Body = body,
                IsBodyHtml = false
            };

            message.To.Add(recipientEmail);

            if (attachment != null && attachment.Length > 0)
            {
                var attachmentStream = new MemoryStream(attachment);
                message.Attachments.Add(new Attachment(attachmentStream, attachmentName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            }
            await smtpClient.SendMailAsync(message);
        }

        // Новый метод для отправки отчётов из Redmine
        public async Task SendRedmineReportAsync(byte[] reportBytes, string nameOfFile, string context)
        {
            // Получаем информацию о пользователе
            // Используем напрямую weldingRedmineService - ключ API уже задан при его создании
            var userInfo = await _weldingRedmineService.GetCurrentUserInfoAsync();
            if (string.IsNullOrEmpty(userInfo?.User?.Mail))
            {
                throw new InvalidOperationException("Не удалось получить email пользователя из Redmine.");
            }

            // Используем существующий метод для отправки
            await SendReportAsync(
                userInfo.User.Mail,
                "Отчёт по сварке",
                "Прикреплённый отчёт во вложении.",
                reportBytes,
                $"{nameOfFile}.xlsx"
            );
        }
    }
}
