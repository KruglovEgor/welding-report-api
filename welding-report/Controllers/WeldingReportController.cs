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

    private readonly RedmineSettings _redmineSettings;

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

    [HttpPost("generate-issue-from-request")]
    public async Task<IActionResult> GenerateIssueFromRequest(
        [FromForm] int issueId = 45,
        [FromForm] string apiKey = "secret"
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

            // ��������� ���������
            var docBytes = _wordReportGenerator.GenerateRequestReport(reportData);

            // ���������� �����
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
            _logger.LogError(ex, "������ ��������� ������");
            return StatusCode(500, ex.Message);
        }
    }

    [HttpPost("generate-issue-from-welding")]
    public async Task<IActionResult> GenerateIssueFromRedmine(
        [FromForm] int issueId = 6,
        [FromForm] string projectName = "test_project",
        [FromForm] string apiKey = "secret",
        [FromForm] bool sendMail = false)
    {
        try
        {
            _redmineService.SetApiKey(apiKey);
            _excelGenerator.SetApiKey(apiKey);
            var reportData = await _redmineService.GetReportDataAsync(projectName, issueId);
            var excelBytes = await _excelGenerator.GenerateIssueReport(reportData);

            // ���������� ������
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
            }


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

    [HttpPost("generate-project-from-welding")]
    public async Task<IActionResult> GenerateProjectFromWelding(
        [FromForm] string projectIdentifier = "test_project",
        [FromForm] string apiKey = "secret",
        [FromForm] bool sendMail = false)
    {
        try
        {
            _redmineService.SetApiKey(apiKey);
            _excelGenerator.SetApiKey(apiKey);
            var projectData = await _redmineService.GetProjectReportDataAsync(projectIdentifier);
            var excelBytes = await _excelGenerator.GenerateProjectReport(projectData);

            // ���������� ������
            var reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{projectData.Name}.xlsx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                await _emailService.SendRedmineReportAsync(excelBytes, projectData.Name,  apiKey, "welding");
            }

            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{projectData.Name}.xlsx"
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��������� ������");
            return StatusCode(500, ex.Message);
        }
    }
}