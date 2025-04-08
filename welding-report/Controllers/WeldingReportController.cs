using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
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
    private readonly INumberToText _numberToText;
    private readonly IRequestWordReportGenerator _wordReportGenerator;

    public WeldingReportController(
        IWebHostEnvironment env,
        ILogger<WeldingReportController> logger,
        IExcelReportGenerator excelGenerator,
        IEmailService emailService,
        IOptions<AppSettings> appSettings,
        IRedmineService redmineService,
        INumberToText numberToText,
        IRequestWordReportGenerator wordReportGenerator
        )
    {
        _env = env;
        _logger = logger;
        _excelGenerator = excelGenerator;
        _emailService = emailService;
        _appSettings = appSettings.Value;
        _redmineService = redmineService;
        _numberToText = numberToText;
        _wordReportGenerator = wordReportGenerator;
    }

    [HttpGet("generate-issue-from-request")]
    public async Task<IActionResult> GenerateIssueFromRequest(
        [FromQuery] int issueId = 45,
        [FromQuery] string apiKey = "secret"
        //[FromForm] string projectName = "portal_zayavok_2"
        )
    {
        try
        {
            _redmineService.SetApiKey(apiKey);
            var reportData = await _redmineService.GetRequestReportDataAsync(issueId);
            _numberToText.FillCostText(reportData);

            if (string.IsNullOrEmpty(reportData.CuratorEmail))
            {
                _logger.LogWarning("CuratorEmail is empty for issue {IssueId}", issueId);
            }

            // Генерация документа
            var docBytes = _wordReportGenerator.GenerateRequestReport(reportData);

            // Сохранение файла
            var reportPath = Path.Combine(
                _env.ContentRootPath,
               _appSettings.ReportStoragePath,
                $"{reportData.Name}.docx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, docBytes);

            return File(
            docBytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            $"{reportData.Name}.docx"
        );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");
            return StatusCode(500, ex.Message);
        }
    }

    [HttpGet("generate-issue-from-welding")]
    public async Task<IActionResult> GenerateIssueFromWelding(
        [FromQuery] int issueId = 6,
        [FromQuery] string projectName = "test_project",
        [FromQuery] string apiKey = "secret",
        [FromQuery] bool sendMail = false)
    {
        try
        {
            _redmineService.SetApiKey(apiKey);
            _excelGenerator.SetApiKey(apiKey);
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
                await _emailService.SendRedmineReportAsync(excelBytes, reportData.ReportNumber, apiKey, "welding");
                return Ok("Отчет успешно создан и отправлен по электронной почте");
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

    [HttpGet("generate-project-from-welding")]
    public async Task<IActionResult> GenerateProjectFromWelding(
        [FromQuery] string projectIdentifier = "test_project",
        [FromQuery] string apiKey = "secret",
        [FromQuery] bool sendMail = false)
    {
        try
        {
            _redmineService.SetApiKey(apiKey);
            _excelGenerator.SetApiKey(apiKey);
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
                await _emailService.SendRedmineReportAsync(excelBytes, projectData.Name, apiKey, "welding");
                return Ok("Отчет успешно создан и отправлен по электронной почте");
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
}