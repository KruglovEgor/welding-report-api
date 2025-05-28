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
    )
    {
        string reportPath = string.Empty;
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
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportData.Name}.docx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, docBytes);

            // ���������� ���� � ��������� ������ �� �������� ����� ��������
            var result = File(
                docBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"{reportData.Name}.docx"
            );

            // ��������� �������� ����� ����� �������� ����������
            _ = Task.Run(async () => {
                // ��������� �������� ��� ���������� �������� �����
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��������� ������");

            // ������� ���� � ������ ������, ���� �� ��� ������
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }


    [HttpGet("generate-issue-from-welding")]
    public async Task<IActionResult> GenerateIssueFromWelding(
        [FromQuery] int issueId = 6,
        [FromQuery] int projectIdentifier = 1,
        [FromQuery] string apiKey = "secret",
        [FromQuery] bool sendMail = false)
    {
        string reportPath = string.Empty;
        try
        {
            _redmineService.SetApiKey(apiKey);
            _excelGenerator.SetApiKey(apiKey);
            var reportData = await _redmineService.GetWeldingIssueDataAsync(projectIdentifier, issueId);
            var excelBytes = await _excelGenerator.GenerateIssueReport(reportData);

            // ���������� ������
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportData.ReportNumber}.xlsx"
            );
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                await _emailService.SendRedmineReportAsync(excelBytes, reportData.ReportNumber, apiKey, "welding");

                // �������� ����� ����� �������� email
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }

                return Ok("����� ������� ������ � ��������� �� ����������� �����");
            }

            // ���������� ���� � ��������� ������ �� �������� ����� ��������
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{reportData.ReportNumber}.xlsx"
            );

            // ��������� �������� ����� ����� �������� ����������
            _ = Task.Run(async () => {
                // ��������� �������� ��� ���������� �������� �����
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��������� ������");

            // ������� ���� � ������ ������, ���� �� ��� ������
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }

    [HttpGet("generate-project-from-welding")]
    public async Task<IActionResult> GenerateProjectFromWelding(
        [FromQuery] int projectIdentifier = 1,
        [FromQuery] string apiKey = "secret",
        [FromQuery] bool sendMail = false)
    {
        string reportPath = string.Empty;
        try
        {
            _redmineService.SetApiKey(apiKey);
            _excelGenerator.SetApiKey(apiKey);
            var projectData = await _redmineService.GetProjectReportDataAsync(projectIdentifier);
            var excelBytes = await _excelGenerator.GenerateProjectReport(projectData);

            // ���������� ������
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{projectData.Name}.xlsx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                await _emailService.SendRedmineReportAsync(excelBytes, projectData.Name, apiKey, "welding");

                // �������� ����� ����� �������� email
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }

                return Ok("����� ������� ������ � ��������� �� ����������� �����");
            }

            // ���������� ���� � ��������� ������ �� �������� ����� ��������
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{projectData.Name}.xlsx"
            );

            // ��������� �������� ����� ����� �������� ����������
            _ = Task.Run(async () => {
                // ��������� �������� ��� ���������� �������� �����
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��������� ������");

            // ������� ���� � ������ ������, ���� �� ��� ������
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }


    [HttpGet("generate-group-from-supr")]
    public async Task<IActionResult> GenerateGroupFromSupr(
        [FromQuery] string projectIdentifier = "test_project",
        [FromQuery] string apiKey = "c022be329bc45f078b14c50b95b0bf4177257c75",
        [FromQuery] int ApplicationNumber = 222
        )
    {
        string reportPath = string.Empty;
        try
        {
            _redmineService.SetApiKey(apiKey);
            var projectData = await _redmineService.GetProjectReportDataAsync(projectIdentifier);
            var excelBytes = await _excelGenerator.GenerateProjectReport(projectData);

            // ���������� ������
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{projectData.Name}.xlsx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                await _emailService.SendRedmineReportAsync(excelBytes, projectData.Name, apiKey, "welding");

                // �������� ����� ����� �������� email
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }

                return Ok("����� ������� ������ � ��������� �� ����������� �����");
            }

            // ���������� ���� � ��������� ������ �� �������� ����� ��������
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{projectData.Name}.xlsx"
            );

            // ��������� �������� ����� ����� �������� ����������
            _ = Task.Run(async () => {
                // ��������� �������� ��� ���������� �������� �����
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "������ ��������� ������");

            // ������� ���� � ������ ������, ���� �� ��� ������
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }
}