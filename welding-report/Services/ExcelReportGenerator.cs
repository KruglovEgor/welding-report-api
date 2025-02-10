namespace welding_report.Services
{
    using OfficeOpenXml;
    using welding_report.Models;
    using SkiaSharp;
    using System.Drawing;
    using Microsoft.Extensions.Options;

    public interface IExcelReportGenerator
    {
        byte[] GenerateReport(WeldingReportRequest request, Dictionary<string, string> photoMap);
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

        public byte[] GenerateReport(WeldingReportRequest request, Dictionary<string, string> photoMap)
        {
            var templateFile = new FileInfo(_templatePath);

            if (!templateFile.Exists)
                throw new FileNotFoundException("Шаблон отчета не найден", _templatePath);

            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[_worksheetName];
            FillData(worksheet, request, photoMap);
            return package.GetAsByteArray();
        }

        private void FillData(ExcelWorksheet worksheet, WeldingReportRequest request, Dictionary<string, string> photoMap)
        {
            const int startRow = 3;
            const int photoColumn = 9;
            _logger.LogInformation($"Total joints: {request.Joints.Count}");

            worksheet.Cells[startRow - 1, 1].Value += " " + request.ReportNumber;

            for (var i = 0; i < request.Joints.Count; i++)
            {
                var joint = request.Joints[i];
                var row = startRow + i;

                worksheet.Cells[row, 1].Value = request.ReportNumber;
                worksheet.Cells[row, 2].Value = string.IsNullOrEmpty(joint.EquipmentType) ? "-" : joint.EquipmentType;
                worksheet.Cells[row, 3].Value = joint.PipelineNumber;
                worksheet.Cells[row, 4].Value = joint.CompanyName;
                worksheet.Cells[row, 5].Value = joint.JointNumber;
                worksheet.Cells[row, 6].Value = joint.DiameterMm > 0 ? joint.DiameterMm : "-";
                worksheet.Cells[row, 7].Value = joint.LengthMeters > 0 ? joint.LengthMeters : "-";
                worksheet.Cells[row, 8].Value = joint.DiameterMm > 0 ? joint.DiameterMm / 25.4 : "-";

                // Устанавливаем фон для второй колонки (EquipmentType)
                var equipmentCell = worksheet.Cells[row, 2];
                equipmentCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                equipmentCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(214, 220, 228));

                // Добавление всех фото
                double maxHeight = 50; // Высота будет хотя бы 50
                double currentWidth = 5; // Отслеживает ширину занятого пространства

                foreach (var photoName in joint.PhotoNames)
                {
                    if (photoMap.TryGetValue(photoName, out var photoPath))
                    {
                        var imgSize = InsertImage(worksheet, row, photoColumn, photoPath, currentWidth);
                        currentWidth += imgSize.width + 5; // Добавляем небольшой отступ
                        maxHeight = Math.Max(maxHeight, imgSize.height);
                    }
                }

                worksheet.Row(row).Height = Math.Min(_maxRowHeight, maxHeight);
                worksheet.Column(photoColumn).Width = Math.Max(currentWidth / 7.0, worksheet.Column(photoColumn).Width); // Excel использует особую шкалу ширины
                var range = worksheet.Cells[row, 1, row, photoColumn];

                // Устанавливаем границы для каждой ячейки в строке
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }
        }

        private (double width, double height) InsertImage(ExcelWorksheet worksheet, int row, int column, string imagePath, double xOffset)
        {
            using var imageStream = File.OpenRead(imagePath);
            using var bitmap = SKBitmap.Decode(imageStream);

            if (bitmap == null)
                throw new Exception($"Не удалось загрузить изображение: {imagePath}");

            double scale = Math.Min(1.0, _maxRowHeight / (double)bitmap.Height);
            int newWidth = (int)(bitmap.Width * scale);
            int newHeight = (int)(bitmap.Height * scale);

            var picture = worksheet.Drawings.AddPicture($"Photo_{row}_{column}_{xOffset}", imagePath);
            picture.SetPosition(row - 1, 5, column - 1, (int)xOffset);
            picture.SetSize(newWidth, newHeight);

            return (newWidth, newHeight);
        }
    }
}
