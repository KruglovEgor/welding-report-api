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
    private const string UploadsFolder = "uploads";
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
            // ��������� email
            if (!IsValidEmail(recipientEmail))
                return BadRequest("�������� ������ email.");

            // ����� ������
            var reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportNumber}.xlsx"
            );

            if (!System.IO.File.Exists(reportPath))
                return NotFound("����� �� ������.");

            var attachment = await System.IO.File.ReadAllBytesAsync(reportPath);
            await _emailService.SendReportAsync(recipientEmail, "����� �� ������", "������������� ����� �� ��������.", attachment, $"{reportNumber}.xlsx");

            return Ok("����� ��������� �� email.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��� �������� ������");
            return StatusCode(500, "������ �������");
        }
    }

    [HttpPost("generate-issue-from-redmine")]
    public async Task<IActionResult> GenerateIssueFromRedmine(
        [FromForm] int issueId,
        [FromForm] string projectName)
    {
        try
        {
            var reportData = await _redmineService.GetReportDataAsync(projectName, issueId);
            var excelBytes = await _excelGenerator.GenerateReport(reportData);

            // ���������� ������
            var reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportData.ReportNumber}.xlsx"
            );
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{reportData.ReportNumber}.xlsx"
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��������� ������");
            return StatusCode(500, ex.Message);
        }
    }

    //[HttpPost("generate-project-from-redmine/{issueId}")]
    //public async Task<IActionResult> GenerateProjectFromRedmine(int issueId)
    //{
    //    try
    //    {
    //        var reportData = await _redmineService.GetReportDataAsync(issueId);
    //        var excelBytes = await _excelGenerator.GenerateReport(reportData);

    //        // ���������� ������
    //        var reportPath = Path.Combine(
    //            _env.ContentRootPath,
    //            _appSettings.ReportStoragePath,
    //            $"{reportData.ReportNumber}.xlsx"
    //        );
    //        Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
    //        await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

    //        return File(
    //            excelBytes,
    //            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    //            $"{reportData.ReportNumber}.xlsx"
    //        );
    //    }
    //    catch (Exception ex)
    //    {
    //        _logger.LogError(ex, "������ ��������� ������");
    //        return StatusCode(500, ex.Message);
    //    }
    //}

    private bool IsValidEmail(string email)
    {
        return Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
    }
}