namespace welding_report.Services
{
    using OfficeOpenXml;
    using welding_report.Models;
    using SkiaSharp;
    using System.Drawing;
    using Microsoft.Extensions.Options;
    using static System.Runtime.InteropServices.JavaScript.JSType;
    using System.Net;
    using OfficeOpenXml.Style;
    using System.Data.Common;

    public interface IExcelReportGenerator
    {
        Task<byte[]> GenerateReport(RedmineReportData data);
    }

    public class ExcelReportGenerator : IExcelReportGenerator
    {
        private readonly string _templatePath;
        private readonly string _worksheetName;
        private readonly int _maxRowHeight;
        private readonly ILogger<ExcelReportGenerator> _logger;

        public ExcelReportGenerator(
            ILogger<ExcelReportGenerator> logger,
            IOptions<AppSettings> appSettings)
        {
            _logger = logger;
            _templatePath = appSettings.Value.TemplatePath;
            _worksheetName = appSettings.Value.WorksheetName;
            _maxRowHeight = appSettings.Value.MaxRowHeight;
        }

        public async Task<byte[]> GenerateReport(RedmineReportData data)
        {
            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[_worksheetName];

            int row = 3;
            foreach (var group in data.Groups)
            {
                var groupStartRow = row;

                foreach (var entry in group.Entries)
                {
                    // Общие данные группы
                    worksheet.Cells[row, 1].Value = data.ReportNumber;
                    worksheet.Cells[row, 6].Value = group.DiameterMm;
                    worksheet.Cells[row, 8].Value = group.DiameterInches;

                    // Данные подрядчика
                    worksheet.Cells[row, 4].Value = entry.Contractor;
                    worksheet.Cells[row, 5].Value = entry.JointNumbers;

                    // Вставка фото
                    double currentWidth = 5;
                    double maxHeight = 50;
                    int photoColumn = 9;

                    foreach (var photoUrl in entry.PhotoUrls)
                    {
                        try
                        {
                            _logger.LogInformation($"{entry.JointNumbers.ToString()} - {photoUrl}");
                            //using var webClient = new WebClient();
                            //byte[] imageBytes = webClient.DownloadData(photoUrl);

                            using var httpClient = new HttpClient();
                            var response = await httpClient.GetAsync(photoUrl);
                            if (!response.IsSuccessStatusCode)
                            {
                                _logger.LogError($"Ошибка HTTP: {response.StatusCode} для {photoUrl}");
                                continue;
                            }

                            // Проверка MIME-типа
                            if (!response.Content.Headers.ContentType.MediaType.StartsWith("image/"))
                            {
                                _logger.LogError($"Недопустимый MIME-тип: {response.Content.Headers.ContentType}");
                                continue;
                            }

                            byte[] imageBytes = await response.Content.ReadAsByteArrayAsync();

                            using var imageStream = new MemoryStream(imageBytes);
                            var imgSize = InsertImage(worksheet, row, photoColumn, imageStream, currentWidth);
                            _logger.LogInformation($"Size: {imgSize}");

                            currentWidth += imgSize.width + 5;
                            maxHeight = Math.Max(maxHeight, imgSize.height);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"Ошибка загрузки фото: {photoUrl}\n{ex.Message}");
                        }
                    }

                    // Настройка размеров строки и столбца
                    worksheet.Row(row).Height = Math.Min(_maxRowHeight, maxHeight);
                    worksheet.Column(photoColumn).Width = Math.Max(currentWidth / 7.0, worksheet.Column(photoColumn).Width);

                    // Границы ячеек
                    var range = worksheet.Cells[row, 1, row, photoColumn];
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    row++;
                }

                // Объединение ячеек группы
                if (row > groupStartRow)
                {
                    worksheet.Cells[groupStartRow, 1, row - 1, 1].Merge = true;
                    worksheet.Cells[groupStartRow, 6, row - 1, 6].Merge = true;
                    worksheet.Cells[groupStartRow, 8, row - 1, 8].Merge = true;
                }
            }

            return package.GetAsByteArray();
        }



        // Метод для расчёта размеров
        private (double width, double height) InsertImage(
            ExcelWorksheet worksheet,
            int row,
            int column,
            Stream imageStream,
            double xOffset)
        {
            // Создаем копию потока для расчета размеров
            byte[] imageBytes;
            using (var tempStream = new MemoryStream())
            {
                imageStream.CopyTo(tempStream);
                imageBytes = tempStream.ToArray();
            }

            // Расчет размеров изображения
            using (var sizeStream = new MemoryStream(imageBytes))
            using (var bitmap = SKBitmap.Decode(sizeStream))
            {
                if (bitmap == null)
                    throw new Exception("Неверный формат изображения");

                if (bitmap.Height == 0 || bitmap.Width == 0)
                {
                    _logger.LogError($"Нулевые размеры изображения: {bitmap.Width}x{bitmap.Height}");
                    return (0, 0);
                }

                _logger.LogInformation($"Bitmap: {bitmap.Width}x{bitmap.Height}");
                
                double scale = Math.Min(1.0, _maxRowHeight / bitmap.Height);
                _logger.LogInformation($"MRH: {_maxRowHeight}");
                int newWidth = (int)(bitmap.Width * scale);
                int newHeight = (int)(bitmap.Height * scale);

                _logger.LogInformation($"New Bitmap: {newWidth}x{newHeight}");

                // Вставка изображения из нового потока
                using (var insertStream = new MemoryStream(imageBytes))
                {
                    var picture = worksheet.Drawings.AddPicture(
                        $"Photo_{Guid.NewGuid()}",
                        insertStream
                    );

                    // Упрощенное позиционирование (в пикселях)
                    picture.SetPosition(
                        row - 1,      // Строка
                        5,            // Смещение Y (в пикселях)
                        column - 1,   // Колонка
                        (int)xOffset  // Смещение X (в пикселях)
                    );

                    picture.SetSize(newWidth, newHeight);
                }

                return (newWidth, newHeight);
            }
        }
    }
}
