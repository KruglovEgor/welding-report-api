using Microsoft.Extensions.Options;
using System.Net.Mail;
using System.Net;
using welding_report.Models;

namespace welding_report.Services
{
    public interface IEmailService
    {
        Task SendReportAsync(string recipientEmail, string subject, string body, byte[] attachment, string attachmentName);
        Task SendRedmineReportAsync(byte[] reportBytes, string reportNumber); 
    }


    public class EmailService : IEmailService
    {
        private readonly EmailSettings _emailSettings;
        private readonly IRedmineService _redmineService;

        public EmailService(
            IOptions<EmailSettings> emailSettings,
            IRedmineService redmineService)
        {
            _emailSettings = emailSettings.Value;
            _redmineService = redmineService;
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
        public async Task SendRedmineReportAsync(byte[] reportBytes, string reportNumber)
        {
            // Получаем email пользователя из Redmine
            var userInfo = await _redmineService.GetCurrentUserInfoAsync();
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
                $"{reportNumber}.xlsx"
            );
        }
    }
}
