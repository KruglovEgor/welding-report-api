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

    [HttpGet("redmine-test/{issueId}")]
    [SwaggerOperation(Summary = "���� ����������� � ��������� ������ �� Redmine")]
    public async Task<IActionResult> TestRedmineAuth(int issueId)
    {
        try
        {
            // ��������� ������������� ����
            var parentResponse = await _redmineService.GetIssueAsync<dynamic>(issueId);
            if (parentResponse?.GetProperty("issue").GetProperty("id").GetInt32() != issueId)
                return NotFound("��� �� ������");

            //if (parentResponse?.Issue == null)
            //    return NotFound("��� �� ������");


            // ��������� �������� ����� ������
            var childrenResponse = await _redmineService.GetChildIssuesAsync<dynamic>(issueId);
            var children = new List<object>();

            foreach (var child in childrenResponse?.GetProperty("issues").EnumerateArray())
            {
                children.Add(child);
            }


            // ������������ ������
            var result = new
            {
                Parent = parentResponse,
                Children = children
            };

            return Ok(result);

        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "������ ������� � Redmine");
            return StatusCode(500, $"������: {ex.Message}");
        }
    }

    [HttpPost("generate-from-redmine/{issueId}")]
    public async Task<IActionResult> GenerateFromRedmine(int issueId)
    {
        try
        {
            var reportData = await _redmineService.GetReportDataAsync(issueId);
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

    private bool IsValidEmail(string email)
    {
        return Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
    }

    private async Task<Dictionary<string, string>> SavePhotos(List<IFormFile> photos)
    {
        var uploadsPath = Path.Combine(_env.ContentRootPath, UploadsFolder);
        Directory.CreateDirectory(uploadsPath);

        var photoMap = new Dictionary<string, string>();

        foreach (var photo in photos)
        {
            if (photo.Length == 0) continue;

            var safeFileName = $"{Guid.NewGuid()}{Path.GetExtension(photo.FileName)}";
            var filePath = Path.Combine(uploadsPath, safeFileName);

            await using var stream = new FileStream(filePath, FileMode.Create);
            await photo.CopyToAsync(stream);

            photoMap[photo.FileName] = filePath;
        }

        return photoMap;
    }

    // �������� ������ ��� docker. � ������� - ���
    private void CleanupFiles(Dictionary<string, string> photoMap)
    {
        foreach (var path in photoMap.Values)
            {
                try
                {
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path);
                        _logger.LogInformation($"Deleted temporary file: {path}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error deleting file: {path}");
                }
            }
        
    }
}