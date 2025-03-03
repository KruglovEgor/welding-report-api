using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Swashbuckle.AspNetCore.Annotations;
using welding_report.Models;
using welding_report.Services;

[ApiController]
[Route("api/[controller]")]
public class WeldingReportController : ControllerBase
{
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<WeldingReportController> _logger;
    private readonly IExcelReportGenerator _excelGenerator;
    private readonly IEmailService _emailService;
    private readonly AppSettings _appSettings;
    private readonly IRedmineService _redmineService;

    public WeldingReportController(
        IWebHostEnvironment env,
        ILogger<WeldingReportController> logger,
        IExcelReportGenerator excelGenerator,
        IEmailService emailService,
        IOptions<AppSettings> appSettings,
        IRedmineService redmineService)
    {
        _env = env;
        _logger = logger;
        _excelGenerator = excelGenerator;
        _emailService = emailService;
        _appSettings = appSettings.Value;
        _redmineService = redmineService;
    }

    [HttpPost("send-report")]
    public async Task<IActionResult> SendReport(
        [FromForm] string recipientEmail,
        [FromForm] string reportNumber)
    {
        try
    {
            // Валидация email
            if (!IsValidEmail(recipientEmail))
                return BadRequest("Неверный формат email.");

            // Поиск отчёта
            var reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportNumber}.xlsx"
            );

            if (!System.IO.File.Exists(reportPath))
                return NotFound("Отчёт не найден.");

            var attachment = await System.IO.File.ReadAllBytesAsync(reportPath);
            await _emailService.SendReportAsync(recipientEmail, "Отчет по сварке", "Прикрепленный отчет во вложении.", attachment, $"{reportNumber}.xlsx");

            return Ok("Отчет отправлен на email.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка при отправке отчета");
            return StatusCode(500, "Ошибка сервера");
        }
    }

    [HttpPost("generate-issue-from-redmine")]
    public async Task<IActionResult> GenerateIssueFromRedmine(
        [FromForm] int issueId = 6,
        [FromForm] string projectName = "test_project",
        [FromForm] bool sendMail = false)
    {
        try
        {
            var reportData = await _redmineService.GetReportDataAsync(projectName, issueId);
            var excelBytes = await _excelGenerator.GenerateIssueReport(reportData);

            // Сохранение отчета
            var reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportData.ReportNumber}.xlsx"
            );
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                await _emailService.SendRedmineReportAsync(excelBytes, reportData.ReportNumber);
            }


            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{reportData.ReportNumber}.xlsx"
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");
            return StatusCode(500, ex.Message);
        }
    }

    [HttpPost("generate-project-from-redmine")]
    public async Task<IActionResult> GenerateProjectFromRedmine(
        [FromForm] string projectIdentifier = "test_project",
        [FromForm] bool sendMail = false)
    {
        try
        {
            var projectData = await _redmineService.GetProjectReportDataAsync(projectIdentifier);
            var excelBytes = await _excelGenerator.GenerateProjectReport(projectData);

            // Сохранение отчета
            var reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{projectData.Name}.xlsx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                await _emailService.SendRedmineReportAsync(excelBytes, projectData.Name);
            }

            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{projectData.Name}.xlsx"
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");
            return StatusCode(500, ex.Message);
        }
    }

    private bool IsValidEmail(string email)
    {
        return Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
    }
}