using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;
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

    public WeldingReportController(
        IWebHostEnvironment env,
        ILogger<WeldingReportController> logger,
        IExcelReportGenerator excelGenerator,
        IEmailService emailService)
    {
        _env = env;
        _logger = logger;
        _excelGenerator = excelGenerator;
        _emailService = emailService;
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
    public async Task<IActionResult> SendReport([FromForm] string recipientEmail)
    {
        try
        {
            var filePath = Path.Combine("Resources/Templates/WeldingReportTemplate.xlsx");
            _logger.LogInformation(filePath);
            if (!System.IO.File.Exists(filePath))
                return NotFound("Файл отчета не найден.");

            var attachment = await System.IO.File.ReadAllBytesAsync(filePath);
            await _emailService.SendReportAsync(recipientEmail, "Отчет по сварке", "Прикрепленный отчет во вложении.", attachment, "Report.xlsx");

            return Ok("Отчет отправлен на email.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка при отправке отчета");
            return StatusCode(500, "Ошибка сервера");
        }
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