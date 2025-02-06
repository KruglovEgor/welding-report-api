using Microsoft.Extensions.Options;
using System.Net.Mail;
using System.Net;
using welding_report.Models;

namespace welding_report.Services
{
    public interface IEmailService
    {
        Task SendReportAsync(string recipientEmail, string subject, string body, byte[] attachment, string attachmentName);
    }


    public class EmailService : IEmailService
    {
        private readonly EmailSettings _emailSettings;

        public EmailService(IOptions<EmailSettings> emailSettings)
        {
            _emailSettings = emailSettings.Value;
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
    }


}
