namespace welding_report.Services.Welding
{
    using OfficeOpenXml;
    using welding_report.Models;
    using SkiaSharp;
    using Microsoft.Extensions.Options;
    using OfficeOpenXml.Style;
    using System;
    using System.Net;
    using DocumentFormat.OpenXml.Bibliography;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Drawing;
    using OfficeOpenXml.Drawing;

    public interface IWeldingExcelReportGenerator
    {
        Task<byte[]> GenerateIssueReport(WeldingIssueReportData data);
        Task<byte[]> GenerateProjectReport(WeldingProjectReportData data);

    }

    public class WeldingExcelReportGenerator : IWeldingExcelReportGenerator
    {
        private readonly string _templatePath;
        private readonly string _worksheetName;
        private readonly int _maxRowHeight;
        private readonly ILogger<WeldingExcelReportGenerator> _logger;
        private readonly int _maxPhotoColumnWidth;
        private readonly System.Drawing.Color _rowColor1;
        private readonly System.Drawing.Color _rowColor2;
        private readonly int _maxPhotoWidthPx;
        private readonly int _maxPhotoHeightPx;
        private readonly int _photoJpegQuality;
        private readonly string _photoCachePath;

        private readonly int startRow = 2;
        private readonly int photoColumn = 10;
        private readonly int xGap = 5;
        private readonly int yGap = 5;
        private WebClient webClient = new();

        private string _apiKey;


        public WeldingExcelReportGenerator(
           ILogger<WeldingExcelReportGenerator> logger,
           IOptions<AppSettings> appSettings,
           string apiKey
           )
        {
            _logger = logger;
            _templatePath = System.IO.Path.Combine( // Use System.IO.Path explicitly to avoid ambiguity.
                appSettings.Value.TemplatePath,
                "WeldingReportTemplate.xlsx"
                );
            _worksheetName = appSettings.Value.WorksheetName;
            _maxRowHeight = appSettings.Value.MaxRowHeight;
            _maxPhotoColumnWidth = appSettings.Value.MaxPhotoColumnWidth;
            _rowColor1 = ColorTranslator.FromHtml(appSettings.Value.ProjectReportRowColor1 ?? "#F5F5F5");
            _rowColor2 = ColorTranslator.FromHtml(appSettings.Value.ProjectReportRowColor2 ?? "#E8F0FE");
            _maxPhotoWidthPx = appSettings.Value.MaxPhotoWidthPx;
            _maxPhotoHeightPx = appSettings.Value.MaxPhotoHeightPx;
            _photoJpegQuality = appSettings.Value.PhotoJpegQuality;
            _photoCachePath = appSettings.Value.WeldingPhotoCachePath ?? "Resources/Welding/Photos";

            webClient.Headers.Add("X-Redmine-API-Key", apiKey);
        }

        public async Task<byte[]> GenerateIssueReport(WeldingIssueReportData data)
        {
            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[_worksheetName];
            worksheet.Column(photoColumn).Width = _maxPhotoColumnWidth;


            InsertResultLine(worksheet, data, startRow);
            int row = startRow + 1;
            row = FillData(worksheet, data, row);
            
            return package.GetAsByteArray();
        }

        public async Task<byte[]> GenerateProjectReport(WeldingProjectReportData data)
        {
            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[_worksheetName];
            worksheet.Column(photoColumn).Width = _maxPhotoColumnWidth;
            int row = startRow;
            bool useFirstColor = true;

            foreach (var IssueData in data.Acts)
            {
                var color = useFirstColor ? _rowColor1 : _rowColor2;
                var firstRow = row;
                InsertResultLine(worksheet, IssueData, row);
                row += 1;
                var lastRow = FillData(worksheet, IssueData, row, data.Identifier);
                row = lastRow;

                ApplyBackgroundColor(worksheet, firstRow, 1, lastRow - 1, photoColumn, color);
                useFirstColor = !useFirstColor;
            }
            return package.GetAsByteArray();
        }


        private bool InsertResultLine(ExcelWorksheet worksheet, WeldingIssueReportData data, int row)
        {
            try
            {
                int calculatedJointsFact = data.DiametrInchesPlan != 0
                    ? (int)Math.Round(data.DiametrInchesFact / data.DiametrInchesPlan * data.JointsCountPlan, MidpointRounding.AwayFromZero)
                    : 0;

                string resultValue = $"Итого по акту: {data.ReportNumber}";
                worksheet.Cells[row, 1].Value = resultValue;
                worksheet.Cells[row, 6].Value = $"{calculatedJointsFact} из {data.JointsCountPlan}";
                worksheet.Cells[row, 9].Value = $"{data.DiametrInchesFact} из {data.DiametrInchesPlan}";

                // Объединение ячеек
                var mergedCells = worksheet.Cells[row, 1, row, 5];
                mergedCells.Merge = true;

                // Включение переноса текста
                mergedCells.Style.WrapText = true;

                // Границы ячеек
                var range = worksheet.Cells[row, 1, row, photoColumn];
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                worksheet.Row(row).CustomHeight = false;
                worksheet.Row(row).Height = CalculateRowHeight(worksheet, row, mergedCells);

                return true;
            } catch (Exception e)
            {
                _logger.LogError($"Error while generating result line in excel: {e.Message}");
                return false;
            }
        }

        private double CalculateRowHeight(ExcelWorksheet worksheet, int row, ExcelRange mergedCells)
        {
            double deafaultRowHeight = 14.4f;

            string mergedValue = mergedCells.Text;
            // Ширина объединенных ячеек в пикселях
            double totalWidth = 0;
            for (int col = mergedCells.Start.Column; col <= mergedCells.End.Column; col++)
            {
                totalWidth += worksheet.Column(col).Width;
            }

            double newHeight = Math.Ceiling(mergedValue.Length / totalWidth) * deafaultRowHeight;

            for (int col = mergedCells.End.Column+1; col <= photoColumn; col++)
            {
                double columnWidth = worksheet.Column(col).Width;
                var cell = worksheet.Cells[row, col];
                string cellValue = cell.Text;

                if (cellValue != null && cellValue.Length > 0)
                {
                    double cellHeight = Math.Ceiling(cellValue.Length / columnWidth) * deafaultRowHeight;
                    newHeight = Math.Max(newHeight, cellHeight);
                }
            }

            return newHeight;
        }



        private int FillData(ExcelWorksheet worksheet, WeldingIssueReportData data, int row, int identifier = 0)
        {
            int innerRow = row;
            try
            {
                foreach (var group in data.Groups)
                {
                    //_logger.LogInformation($"Group: {group.ToString()}");
                    var groupStartRow = innerRow;

                    worksheet.Cells[innerRow, 1].Value = data.ReportNumber;
                    worksheet.Cells[innerRow, 2].Value = group.ActParagraph;
                    worksheet.Cells[innerRow, 3].Value = group.EquipmentType;
                    worksheet.Cells[innerRow, 4].Value = group.PipelineNumber;

                    worksheet.Cells[innerRow, 7].Value = group.DiameterMm;
                    worksheet.Cells[innerRow, 9].Value = group.DiameterInches;

                    // Устанавливаем фон для второй колонки (EquipmentType)
                    var equipmentCell = worksheet.Cells[innerRow, 3];
                    equipmentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    equipmentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(214, 220, 228));

                    foreach (var entry in group.Entries)
                    {
                        // Собираем все стыки и фото в правильном порядке
                        var sortedJoints = string.Join(", ", entry.JointPhotoMap.Keys);
                        var allPhotos = entry.JointPhotoMap.Values.SelectMany(urls => urls).ToList();

                        worksheet.Cells[innerRow, 5].Value = entry.Contractor;
                        worksheet.Cells[innerRow, 6].Value = sortedJoints;


                        // Вставка фото
                        double currentWidth = xGap;
                        double currentHeight = yGap;
                        int rowHeight = _maxRowHeight;

                        foreach (var photoUrl in allPhotos)
                        {
                            try
                            {
                                //_logger.LogInformation($"{sortedJoints} - {photoUrl}");
                                
                                var imgInfo = InsertImage(worksheet, innerRow, photoColumn, photoUrl, currentWidth, currentHeight, identifier);

                                if (imgInfo.isNewRow)
                                {
                                    rowHeight += _maxRowHeight + yGap;
                                    currentHeight += imgInfo.height + yGap;
                                    currentWidth = 5 + imgInfo.width + xGap;
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
                        worksheet.Row(innerRow).Height = rowHeight + yGap;


                        // Границы ячеек
                        var range = worksheet.Cells[innerRow, 1, innerRow, photoColumn];
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        innerRow++;
                    }

                    // Объединение ячеек группы
                    if (innerRow > groupStartRow)
                    {
                        worksheet.Cells[groupStartRow, 1, innerRow - 1, 1].Merge = true;
                        worksheet.Cells[groupStartRow, 2, innerRow - 1, 2].Merge = true;
                        worksheet.Cells[groupStartRow, 3, innerRow - 1, 3].Merge = true;
                        worksheet.Cells[groupStartRow, 4, innerRow - 1, 4].Merge = true;

                        worksheet.Cells[groupStartRow, 7, innerRow - 1, 7].Merge = true;
                        worksheet.Cells[groupStartRow, 8, innerRow - 1, 8].Merge = true;
                        worksheet.Cells[groupStartRow, 9, innerRow - 1, 9].Merge = true;

                    }
                }

                return innerRow;

            } catch (Exception e)
                {
                _logger.LogError($"Error while filling data in excel: {e.Message}");
                return row;
            }
        }

        private void ApplyBackgroundColor(ExcelWorksheet worksheet, int startRow, int startColumn, int lastRow, int lastColumn, System.Drawing.Color color)
        {
            var range = worksheet.Cells[startRow, startColumn, lastRow, lastColumn];
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        // Метод для расчёта размеров
        private (double width, double height, bool isNewRow) InsertImage(
            ExcelWorksheet worksheet,
            int row,
            int column,
            string photoUrl,
            double xOffset,
            double yOffset,
            int identifier = 0)
        {
            _logger.LogInformation($"Processing photo: {photoUrl}");

            ExcelPicture picture = null;
            int newWidth = 0;
            int newHeight = 0;
            bool isNewRow = false;
            string cachePath = BuildPhotoCachePath(identifier, photoUrl);

            bool isCached = false;
            
            if (identifier > 0)
            {
                byte[] cachedPhoto = TryGetPhoto(cachePath);
                if (cachedPhoto != null)
                {
                    using (var bitmapStream = new MemoryStream(cachedPhoto))
                    using (var bitmap = SKBitmap.Decode(bitmapStream))
                    {
                        double scaleExcel = _maxRowHeight * 1.3 / bitmap.Height;
                        newWidth = (int)(bitmap.Width * scaleExcel);
                        newHeight = (int)(bitmap.Height * scaleExcel);
                    }

                    var pictureStream = new MemoryStream(cachedPhoto);
                    picture = worksheet.Drawings.AddPicture(
                        $"Photo_{Guid.NewGuid()}",
                        pictureStream
                    );
                    isCached = true;
                    _logger.LogInformation($"Используется кэшированное фото: {cachePath}");
                }
            }

            if (!isCached)
            {
                byte[] imageBytes = webClient.DownloadData(photoUrl);
                using var imageStream = new MemoryStream(imageBytes);

                // Расчет размеров изображения
                using (var bitmap = SKBitmap.Decode(imageStream))
                {
                    if (bitmap == null)
                        throw new Exception("Неверный формат изображения");

                    if (bitmap.Height == 0 || bitmap.Width == 0)
                    {
                        _logger.LogError($"Нулевые размеры изображения: {bitmap.Width}x{bitmap.Height}");
                        return (0, 0, isNewRow);
                    }

                    // 1. Вычисляем коэффициент масштабирования
                    double scaleX = (double)_maxPhotoWidthPx / bitmap.Width;
                    double scaleY = (double)_maxPhotoHeightPx / bitmap.Height;
                    double scale = Math.Min(1.0, Math.Min(scaleX, scaleY)); // не увеличиваем

                    int resizedWidth = (int)(bitmap.Width * scale);
                    int resizedHeight = (int)(bitmap.Height * scale);

                    using (var compressedStream = new MemoryStream())
                    {
                        using (var resizedBitmap = bitmap.Resize(new SKImageInfo(resizedWidth, resizedHeight), SKFilterQuality.Medium))
                        using (var image = SKImage.FromBitmap(resizedBitmap))
                        using (var data = image.Encode(SKEncodedImageFormat.Jpeg, _photoJpegQuality))
                        {
                            data.SaveTo(compressedStream);
                        }
                        compressedStream.Position = 0;

                        if (identifier > 0)
                        {
                            SavePhoto(cachePath, compressedStream.ToArray());
                        }

                        // 3. Для отображения в Excel используем прежнюю логику масштабирования
                        double scaleExcel = _maxRowHeight * 1.3 / bitmap.Height;
                        newWidth = (int)(bitmap.Width * scaleExcel);
                        newHeight = (int)(bitmap.Height * scaleExcel);

                        picture = worksheet.Drawings.AddPicture(
                            $"Photo_{Guid.NewGuid()}",
                            compressedStream
                        );
                    }
                }
             }

            if ((xOffset + newWidth) / 7.0 > _maxPhotoColumnWidth)
            {
                isNewRow = true;
                yOffset += newHeight + yGap;
                xOffset = xGap;
            }

            // Упрощенное позиционирование (в пикселях)
            picture.SetPosition(
                row - 1,      // Строка
                (int)yOffset, // Смещение Y (в пикселях)
                column - 1,   // Колонка
                (int)xOffset  // Смещение X (в пикселях)
            );
            picture.SetSize(newWidth, newHeight);

            return (newWidth, newHeight, isNewRow);
        }
         

        // Построение пути по URL и identifier
        private string BuildPhotoCachePath(int identifier, string photoUrl)
        {
            // Извлекаем часть после "attachments/download/"
            var uri = new Uri(photoUrl);
            var parts = uri.AbsolutePath.Split(new[] { "attachments/download/" }, StringSplitOptions.None);
            if (parts.Length < 2)
                throw new ArgumentException("Некорректный формат URL для фото");

            var filePart = parts[1].Replace('/', '_');
            var dir = System.IO.Path.Combine(_photoCachePath, identifier.ToString());
            Directory.CreateDirectory(dir);
            return System.IO.Path.Combine(dir, filePart);
        }

        // 1. Получить фото по пути (или null, если нет)
        private byte[] TryGetPhoto(string path)
        {
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }

        // 2. Сохранить фото по пути (перезаписывает)
        private void SavePhoto(string path, byte[] data)
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, data);
        }

    }
}
