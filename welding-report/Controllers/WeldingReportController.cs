using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using welding_report.Models;

[ApiController]
[Route("api/[controller]")]
public class WeldingReportController : ControllerBase
{
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<WeldingReportController> _logger;
    private const string UploadsFolder = "uploads";

    public WeldingReportController(
        IWebHostEnvironment env,
        ILogger<WeldingReportController> logger)
    {
        _env = env;
        _logger = logger;
    }

    [HttpPost("generate")]
    public async Task<IActionResult> GenerateReport([FromForm] WeldingReportRequest request)
    {
        try
        {
            // Валидация
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var uploadPath = Path.Combine(_env.ContentRootPath, "uploads");
            Directory.CreateDirectory(uploadPath);

            var photoMap = await SavePhotos(request.Photos);
            var excelBytes = GenerateExcel(request, photoMap);

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
            //return StatusCode(500, "Internal Server Error");
            return StatusCode(500, ex.Message);
        }
    }

    private byte[] GenerateExcel(WeldingReportRequest request, Dictionary<string, List<string>> photoMap)
    {
        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Отчет");

        // Заголовок
        ws.Cells["A1"].Value = "№ Акта:";
        ws.Cells["B1"].Value = request.ReportNumber;

        // Шапка таблицы
        var headers = new[]
        {
            "Тип оборудования", "№ трубопровода",
            "№ Стыка", "Диаметр (мм)",
            "п.м.", "Дюймаж", "Фото"
        };

        ws.Cells["A3"].LoadFromArrays(new[] { headers });

        // Данные
        var row = 4;
        foreach (var joint in request.Joints)
        {
            ws.Cells[row, 1].Value = joint.EquipmentType;
            ws.Cells[row, 2].Value = joint.PipelineNumber;
            ws.Cells[row, 3].Value = joint.JointNumber;
            ws.Cells[row, 4].Value = joint.DiameterMm;
            ws.Cells[row, 5].Value = joint.LengthMeters;
            ws.Cells[row, 6].Value = joint.DiameterMm / 25.4;

            // Вставка фото
            if (photoMap.TryGetValue(joint.JointNumber, out var photos))
            {
                using var imageStream = new FileStream(photos.First(), FileMode.Open);
                var excelImage = ws.Drawings.AddPicture($"Photo{row}", imageStream);
                excelImage.SetPosition(row - 1, 0, 6, 0);
                excelImage.SetSize(100, 100);
            }

            row++;
        }

        // Форматирование
        ws.Cells[ws.Dimension.Address].AutoFitColumns();
        ws.Cells["D4:D100"].Style.Numberformat.Format = "0.00";
        ws.Cells["E4:E100"].Style.Numberformat.Format = "0.00";
        ws.Cells["F4:F100"].Style.Numberformat.Format = "0.00";

        return package.GetAsByteArray();
    }

    private async Task<Dictionary<string, List<string>>> SavePhotos(List<IFormFile> photos)
    {
        var uploadsPath = Path.Combine(_env.ContentRootPath, "uploads");
        if (!Directory.Exists(uploadsPath))
        {
            Directory.CreateDirectory(uploadsPath);
            _logger.LogInformation($"Created uploads directory at: {uploadsPath}");
        }

        var photoMap = new Dictionary<string, List<string>>();

        foreach (var photo in photos)
        {
            if (photo.Length == 0) continue;

            var jointNumber = Path.GetFileNameWithoutExtension(photo.FileName)?
                .Split('_')[0]?
                .Trim();

            if (string.IsNullOrEmpty(jointNumber))
            {
                _logger.LogWarning($"Invalid photo filename: {photo.FileName}");
                continue;
            }

            var safeFileName = $"{jointNumber}_{Guid.NewGuid()}{Path.GetExtension(photo.FileName)}";
            var filePath = Path.Combine(uploadsPath, safeFileName);

            await using var stream = new FileStream(filePath, FileMode.Create);
            await photo.CopyToAsync(stream);

            if (!photoMap.ContainsKey(jointNumber))
                photoMap[jointNumber] = new List<string>();

            photoMap[jointNumber].Add(filePath);
        }

        return photoMap;
    }

    private void CleanupFiles(Dictionary<string, List<string>> photoMap)
    {
        foreach (var photos in photoMap.Values)
        {
            foreach (var path in photos)
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
}