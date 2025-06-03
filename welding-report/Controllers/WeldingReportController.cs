using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Swashbuckle.AspNetCore.Annotations;
using welding_report.Models;
using welding_report.Services;
using welding_report.Services.Request;
using welding_report.Services.Welding;

[ApiController]
[Route("api/[controller]")]
public class WeldingReportController : ControllerBase
{
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<WeldingReportController> _logger;
    //private readonly IWeldingExcelReportGenerator _weldingExcelGenerator;
    private readonly IEmailServiceFactory _emailServiceFactory;
    private readonly AppSettings _appSettings;
    private readonly IRedmineServiceFactory _redmineServiceFactory;
    private readonly INumberToText _numberToText;
    private readonly IRequestWordReportGenerator _wordReportGenerator;

    public WeldingReportController(
        IWebHostEnvironment env,
        ILogger<WeldingReportController> logger,
        IEmailServiceFactory emailServiceFactory,
        IOptions<AppSettings> appSettings,
        IRedmineServiceFactory redmineServiceFactory,
        INumberToText numberToText,
        IRequestWordReportGenerator wordReportGenerator
        )
    {
        _env = env;
        _logger = logger;
        _emailServiceFactory = emailServiceFactory;
        _appSettings = appSettings.Value;
        _redmineServiceFactory = redmineServiceFactory;
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
            var requestRedmineService = _redmineServiceFactory.CreateRequestService(apiKey);
            var reportData = await requestRedmineService.GetRequestReportDataAsync(issueId);
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
            var weldingRedmineService = _redmineServiceFactory.CreateWeldingService(apiKey);
            var weldingExcelGenerator = _redmineServiceFactory.CreateWeldingExcelGenerator(apiKey);
            var reportData = await weldingRedmineService.GetWeldingIssueDataAsync(projectIdentifier, issueId);
            var excelBytes = await weldingExcelGenerator.GenerateIssueReport(reportData);


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
                var weldingEmailService = _emailServiceFactory.CreateWeldingEmailService(apiKey);
                await weldingEmailService.SendRedmineReportAsync(excelBytes, reportData.ReportNumber, apiKey);

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
            var weldingRedmineService = _redmineServiceFactory.CreateWeldingService(apiKey);
            var weldingExcelGenerator = _redmineServiceFactory.CreateWeldingExcelGenerator(apiKey);
            var projectData = await weldingRedmineService.GetProjectReportDataAsync(projectIdentifier);
            var excelBytes = await weldingExcelGenerator.GenerateProjectReport(projectData);


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
                var weldingEmailService = _emailServiceFactory.CreateWeldingEmailService(apiKey);
                await weldingEmailService.SendRedmineReportAsync(excelBytes, projectData.Name, apiKey);

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
        [FromQuery] string apiKey = "secret",
        [FromQuery] int applicationNumber = 222
        )
    {
        string reportPath = string.Empty;
        try
        {
            var suprRedmineService = _redmineServiceFactory.CreateSuprService(apiKey);
            var projectData = await suprRedmineService.GetSuprGroupReportDataAsync(projectIdentifier, applicationNumber);
            var excelGenerator = _redmineServiceFactory.CreateSuprExcelGenerator();
            var excelBytes = await excelGenerator.GenerateGroupReport(projectData);

            // ����������� ���� ��� ���������� ���������� �����
            string fileName = $"SUPR_{projectData.Factory}_{applicationNumber}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                fileName
            );

            // ������� ����������, ���� �� ����������
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);


            // ���������� ���� ��� ����������
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName
            );

            // ������� ��������� ���� ����� ��������
            _ = Task.Run(async () => {
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