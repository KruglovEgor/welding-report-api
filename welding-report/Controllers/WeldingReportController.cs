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
            _logger.LogInformation("������� ������ �� ��������� ��������� �� ������. IssueId: {IssueId}", issueId);

            _logger.LogInformation("������ ��������� ������� ��� ������ {IssueId}", issueId);
            var requestRedmineService = _redmineServiceFactory.CreateRequestService(apiKey);
            var reportData = await requestRedmineService.GetRequestReportDataAsync(issueId);
            _numberToText.FillCostText(reportData);

            if (string.IsNullOrEmpty(reportData.CuratorEmail))
            {
                _logger.LogWarning("CuratorEmail ����������� ��� ������ {IssueId}", issueId);
            }

            _logger.LogInformation("������ �������� ��������� Word ��� ������ {IssueId}", issueId);
            // ��������� ���������
            var docBytes = _wordReportGenerator.GenerateRequestReport(reportData);
            _logger.LogInformation("�������� Word ������� ������������ ��� ������ {IssueId}", issueId);

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

            _logger.LogInformation("�������� ������� ������������ � ��������� ������� ��� ������ {IssueId}", issueId);
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
            _logger.LogInformation("������� ������ �� ��������� ������ �� ������. IssueId: {IssueId}, ProjectIdentifier: {ProjectIdentifier}, SendMail: {SendMail}", issueId, projectIdentifier, sendMail);
            var weldingRedmineService = _redmineServiceFactory.CreateWeldingService(apiKey);
            var weldingExcelGenerator = _redmineServiceFactory.CreateWeldingExcelGenerator(apiKey);
            _logger.LogInformation("������ ��������� ������� ��� ������ �� ������ {IssueId}", issueId);
            var reportData = await weldingRedmineService.GetWeldingIssueDataAsync(projectIdentifier, issueId);
            _logger.LogInformation("������ �������� Excel ��������� ��� ������ �� ������ {IssueId}", issueId);
            var excelBytes = await weldingExcelGenerator.GenerateIssueReport(reportData);
            _logger.LogInformation("Excel �������� ������� ������������ ��� ������ �� ������ {IssueId}", issueId);


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
                _logger.LogInformation("������ �������� ������ {IssueId} �� ������ �� email", issueId);
                await weldingEmailService.SendRedmineReportAsync(excelBytes, reportData.ReportNumber, apiKey);
                _logger.LogInformation("����� �� ������ {IssueId} ������� ��������� �� email", issueId);

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
            _logger.LogInformation("�������� ������� ������������ � ��������� ������� ��� ������ �� ������ {IssueId}", issueId);
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
            _logger.LogInformation("������� ������ �� ��������� ������ �� ������� ������. ProjectIdentifier: {ProjectIdentifier}, SendMail: {SendMail}", projectIdentifier, sendMail);
            var weldingRedmineService = _redmineServiceFactory.CreateWeldingService(apiKey);
            var weldingExcelGenerator = _redmineServiceFactory.CreateWeldingExcelGenerator(apiKey);
            _logger.LogInformation("������ ��������� ������� ��� ������ �� ������� {ProjectIdentifier}", projectIdentifier);
            var projectData = await weldingRedmineService.GetProjectReportDataAsync(projectIdentifier);
            _logger.LogInformation("������ �������� Excel ��������� ��� ������� {ProjectName}", projectData.Name);
            var excelBytes = await weldingExcelGenerator.GenerateProjectReport(projectData);
            _logger.LogInformation("Excel �������� ������� ������������ ��� ������� {ProjectName}", projectData.Name);

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
                _logger.LogInformation("������ �������� ������ �� ������� {ProjectName} �� email", projectData.Name);
                await weldingEmailService.SendRedmineReportAsync(excelBytes, projectData.Name, apiKey);
                _logger.LogInformation("����� �� ������� {ProjectName} ������� ��������� �� email", projectData.Name);

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
            _logger.LogInformation("�������� ������� ������������ � ��������� ������� ��� ������� {ProjectName}", projectData.Name);

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
            _logger.LogInformation("������� ������ �� ��������� ������ SUPR. ProjectIdentifier: {ProjectIdentifier}, ApplicationNumber: {ApplicationNumber}", projectIdentifier, applicationNumber);
            var suprRedmineService = _redmineServiceFactory.CreateSuprService(apiKey);
            _logger.LogInformation("������ ��������� ������� SUPR ��� ������� {ProjectIdentifier}, ������ {ApplicationNumber}", projectIdentifier, applicationNumber);
            var projectData = await suprRedmineService.GetSuprGroupReportDataAsync(projectIdentifier, applicationNumber);
            var excelGenerator = _redmineServiceFactory.CreateSuprExcelGenerator();
            _logger.LogInformation("������ �������� Excel ��������� SUPR ��� ����������� {Factory}, ������ {ApplicationNumber}", projectData.Factory, applicationNumber);
            var excelBytes = await excelGenerator.GenerateGroupReport(projectData);
            _logger.LogInformation("Excel �������� SUPR ������� ������������ ��� ����������� {Factory}, ������ {ApplicationNumber}", projectData.Factory, applicationNumber);


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
            _logger.LogInformation("�������� SUPR ������� ������������ � ��������� ������� ��� ����������� {Factory}, ������ {ApplicationNumber}", projectData.Factory, applicationNumber);
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