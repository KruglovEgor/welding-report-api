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
        Task<byte[]> GenerateIssueReport(WeldingReportData data);
        Task<byte[]> GenerateProjectReport(WeldingProjectReportData data);
    }

    public class ExcelReportGenerator : IExcelReportGenerator
    {
        private readonly string _templatePath;
        private readonly string _worksheetName;
        private readonly int _maxRowHeight;
        private readonly ILogger<ExcelReportGenerator> _logger;
        private readonly int _maxPhotoColumnWidth;
        private readonly RedmineSettings _settings;

        private readonly int startRow = 2;
        private readonly int photoColumn = 10;
        private readonly int xGap = 5;
        private readonly int yGap = 5;
        private WebClient webClient = new();
        


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
            // Добавляем заголовок с API-ключом (как в HttpClient)
            webClient.Headers.Add("X-Redmine-API-Key", _settings.WeldingApiKey);
        }

        public async Task<byte[]> GenerateIssueReport(WeldingReportData data)
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
            foreach (var IssueData in data.Acts)
            {
                InsertResultLine(worksheet, IssueData, row);
                row += 1;
                row = FillData(worksheet, IssueData, row);
            }
            return package.GetAsByteArray();
        }


        private bool InsertResultLine(ExcelWorksheet worksheet, WeldingReportData data, int row)
        {
            try
            {
                string resultValue = $"Итого по акту: {data.ReportNumber}";
                worksheet.Cells[row, 1].Value = resultValue;
                worksheet.Cells[row, 6].Value = $"{data.JointsCountFact} из {data.JointsCountPlan}";
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



        private int FillData(ExcelWorksheet worksheet, WeldingReportData data, int row)
        {
            int innerRow = row;
            try
            {
                foreach (var group in data.Groups)
                {
                    _logger.LogInformation($"Group: {group.ToString()}");
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
                    equipmentCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(214, 220, 228));

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
                                _logger.LogInformation($"{sortedJoints} - {photoUrl}");

                                byte[] imageBytes = webClient.DownloadData(photoUrl);

                                using var imageStream = new MemoryStream(imageBytes);
                                var imgInfo = InsertImage(worksheet, innerRow, photoColumn, imageStream, currentWidth, currentHeight);

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
                    yOffset += newHeight + yGap;
                    xOffset = xGap;
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
