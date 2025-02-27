namespace welding_report.Services
{
    using OfficeOpenXml;
    using welding_report.Models;
    using SkiaSharp;
    using Microsoft.Extensions.Options;
    using OfficeOpenXml.Style;
    using System.Drawing;
    using System;
    using System.Net;

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
        private readonly RedmineSettings _settings;

        public ExcelReportGenerator(
            ILogger<ExcelReportGenerator> logger,
            IOptions<AppSettings> appSettings,
            IOptions<RedmineSettings> redmineSettings)
        {
            _logger = logger;
            _templatePath = appSettings.Value.TemplatePath;
            _worksheetName = appSettings.Value.WorksheetName;
            _maxRowHeight = appSettings.Value.MaxRowHeight;
            _maxPhotoColumnWidth = appSettings.Value.MaxPhotoColumnWidth;
            _settings = redmineSettings.Value;
        }

        public async Task<byte[]> GenerateReport(RedmineReportData data)
        {
            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[_worksheetName];

            worksheet.Cells[2, 1].Value += " " + data.ReportNumber;
            worksheet.Cells[2, 6].Value = $"{data.JointsCountFact} из {data.JointsCountPlan}";
            worksheet.Cells[2, 9].Value = $"{data.DiametrInchesFact} из {data.DiametrInchesPlan}";

            int row = 3;
            int photoColumn = 10;
            worksheet.Column(photoColumn).Width = _maxPhotoColumnWidth;

            foreach (var group in data.Groups)
            {
                _logger.LogInformation($"Group: {group.ToString()}");
                var groupStartRow = row;

                worksheet.Cells[row, 1].Value = data.ReportNumber;
                worksheet.Cells[row, 2].Value = group.ActParagraph;
                worksheet.Cells[row, 3].Value = group.EquipmentType;
                worksheet.Cells[row, 4].Value = group.PipelineNumber;

                worksheet.Cells[row, 7].Value = group.DiameterMm;
                worksheet.Cells[row, 9].Value = group.DiameterInches;

                // Устанавливаем фон для второй колонки (EquipmentType)
                var equipmentCell = worksheet.Cells[row, 3];
                equipmentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                equipmentCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(214, 220, 228));

                foreach (var entry in group.Entries)
                {
                    // Собираем все стыки и фото в правильном порядке
                    var sortedJoints = string.Join(", ", entry.JointPhotoMap.Keys);
                    var allPhotos = entry.JointPhotoMap.Values.SelectMany(urls => urls).ToList();

                    worksheet.Cells[row, 5].Value = entry.Contractor;
                    worksheet.Cells[row, 6].Value = sortedJoints;
                    

                    // Вставка фото
                    double currentWidth = 5;
                    double currentHeight = 5;
                    int rowHeight = _maxRowHeight;

                    foreach (var photoUrl in allPhotos)
                    {
                        try
                        {
                            _logger.LogInformation($"{sortedJoints} - {photoUrl}");

                            using var webClient = new WebClient();
                            // Добавляем заголовок с API-ключом (как в HttpClient)
                            webClient.Headers.Add("X-Redmine-API-Key", _settings.ApiKey);

                            // Устанавливаем Accept-заголовок для JSON (если требуется)
                            webClient.Headers.Add(HttpRequestHeader.Accept, "application/json");
                            byte[] imageBytes = webClient.DownloadData(photoUrl);

                            using var imageStream = new MemoryStream(imageBytes);
                            var imgInfo = InsertImage(worksheet, row, photoColumn, imageStream, currentWidth, currentHeight);

                            if (imgInfo.isNewRow)
                            {
                                rowHeight += _maxRowHeight + 5;
                                currentHeight += imgInfo.height + 5;
                                currentWidth = 5 + imgInfo.width + 5;
                            }
                            else
                            {
                                currentWidth += imgInfo.width + 5;
                            }
                                
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"Ошибка загрузки фото: {photoUrl}\n{ex.Message}");
                        }
                    }

                    // Настройка размеров строки и столбца
                    worksheet.Row(row).Height = rowHeight+5;
                    

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
                    
                    worksheet.Cells[groupStartRow, 7, row - 1, 7].Merge = true;
                    worksheet.Cells[groupStartRow, 8, row - 1, 8].Merge = true;
                    worksheet.Cells[groupStartRow, 9, row - 1, 9].Merge = true;

                }
            }

            return package.GetAsByteArray();
        }



        // Метод для расчёта размеров
        private (double width, double height, bool isNewRow) InsertImage(
            ExcelWorksheet worksheet,
            int row,
            int column,
            Stream imageStream,
            double xOffset,
            double yOffset)
        {
            bool isNewRow = false;
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
                    return (0, 0, isNewRow);
                }
                
                //1.3 коэффициент для пикселей (1 единица высоты в excel = 1.(3) пикселя)
                double scale = (double)_maxRowHeight*1.3 / bitmap.Height;
                int newWidth = (int)(bitmap.Width * scale);
                int newHeight = (int)(bitmap.Height * scale);

                _logger.LogInformation($"xOf: {xOffset}, newWid: {newWidth}, value: {(xOffset + newWidth) / 7.0}, maxWid: {_maxPhotoColumnWidth}");

                if ((xOffset+newWidth)/7.0 > _maxPhotoColumnWidth)
                {
                    isNewRow = true;
                    yOffset += newHeight + 5;
                    xOffset = 5;
                }

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
                        (int)yOffset, // Смещение Y (в пикселях)
                        column - 1,   // Колонка
                        (int)xOffset  // Смещение X (в пикселях)
                    );

                    picture.SetSize(newWidth, newHeight);
                }

                return (newWidth, newHeight, isNewRow);
            }
        }
    }
}
