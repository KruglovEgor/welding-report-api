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

    [HttpPost("generate")]
    public async Task<IActionResult> GenerateReport(
    [FromForm] string ReportNumber,
    [FromForm] string Joints, // JSON string
    [FromForm] List<IFormFile> Photos)
    {
        try
        {
            if (string.IsNullOrEmpty(Joints))
            {
                return BadRequest("Joints cannot be empty.");
            }

            var joints = JsonSerializer.Deserialize<List<WeldingJoint>>(Joints);
            if (joints == null || joints.Count == 0)
            {
                return BadRequest("Invalid JSON format or empty joints array.");
            }

            var request = new WeldingReportRequest
            {
                ReportNumber = ReportNumber,
                Joints = joints,
                Photos = Photos ?? new List<IFormFile>()
            };

            var photoMap = await SavePhotos(request.Photos);

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var excelBytes = _excelGenerator.GenerateReport(request, photoMap);
            CleanupFiles(photoMap);
            // В методе GenerateReport после генерации excelBytes:
            var reportPath = Path.Combine(_env.ContentRootPath, _appSettings.ReportStoragePath, $"{request.ReportNumber}.xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            return File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{request.ReportNumber}.xlsx"
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating report");
            return StatusCode(500, ex.Message);
        }
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

    [HttpGet("redmine-test/{issueId}")]
    [SwaggerOperation(Summary = "Тест авторизации и получения данных из Redmine")]
    public async Task<IActionResult> TestRedmineAuth(int issueId)
    {
        try
        {
            // Получение родительского акта
            var parentResponse = await _redmineService.GetIssueAsync<dynamic>(issueId);
            if (parentResponse?.GetProperty("issue").GetProperty("id").GetInt32() != issueId)
                return NotFound("Акт не найден");

            //if (parentResponse?.Issue == null)
            //    return NotFound("Акт не найден");


            // Получение дочерних групп стыков
            var childrenResponse = await _redmineService.GetChildIssuesAsync<dynamic>(issueId);
            var children = new List<object>();

            foreach (var child in childrenResponse?.GetProperty("issues").EnumerateArray())
            {
                children.Add(child);
            }


            // Формирование ответа
            var result = new
            {
                Parent = parentResponse,
                Children = children
            };

            return Ok(result);

        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Ошибка запроса к Redmine");
            return StatusCode(500, $"Ошибка: {ex.Message}");
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

    // Работает только для docker. В локалке - нет
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