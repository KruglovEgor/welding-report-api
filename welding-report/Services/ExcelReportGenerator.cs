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
        private readonly int _maxPhotoColumnWidth;

        public ExcelReportGenerator(
            ILogger<ExcelReportGenerator> logger,
            IOptions<AppSettings> appSettings)
        {
            _logger = logger;
            _templatePath = appSettings.Value.TemplatePath;
            _worksheetName = appSettings.Value.WorksheetName;
            _maxRowHeight = appSettings.Value.MaxRowHeight;
            _maxPhotoColumnWidth = appSettings.Value.MaxPhotoColumnWidth;
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
                    worksheet.Cells[row, 1].Value = data.ReportNumber;
                    worksheet.Cells[row, 2].Value = group.ActParagraph;
                    worksheet.Cells[row, 3].Value = group.EquipmentType;
                    worksheet.Cells[row, 4].Value = group.PipelineNumber;
                    worksheet.Cells[row, 5].Value = entry.Contractor;
                    worksheet.Cells[row, 6].Value = entry.JointNumbers;
                    worksheet.Cells[row, 7].Value = group.DiameterMm;
                    worksheet.Cells[row, 9].Value = group.DiameterInches;

                    // Вставка фото
                    double currentWidth = 5;
                    int photoColumn = 10;

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

                            currentWidth += imgSize.width + 5;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"Ошибка загрузки фото: {photoUrl}\n{ex.Message}");
                        }
                    }

                    // Настройка размеров строки и столбца
                    worksheet.Row(row).Height = _maxRowHeight+5;
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
                    worksheet.Cells[groupStartRow, 2, row - 1, 2].Merge = true;
                    worksheet.Cells[groupStartRow, 3, row - 1, 3].Merge = true;
                    worksheet.Cells[groupStartRow, 4, row - 1, 4].Merge = true;
                    
                    worksheet.Cells[groupStartRow, 6, row - 1, 7].Merge = true;
                    worksheet.Cells[groupStartRow, 8, row - 1, 8].Merge = true;
                    worksheet.Cells[groupStartRow, 9, row - 1, 9].Merge = true;

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
                
                //1.3 коэффициент для пикселей (1 единица высоты в excel = 1.(3) пикселя)
                double scale = (double)_maxRowHeight*1.3 / bitmap.Height;
                int newWidth = (int)(bitmap.Width * scale);
                int newHeight = (int)(bitmap.Height * scale);

                //_logger.LogInformation($"New Bitmap: {newWidth}x{newHeight}");

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
