namespace welding_report.Services
{
    using OfficeOpenXml;
    using welding_report.Models;
    using SkiaSharp;
    using System.Drawing;

    public interface IExcelReportGenerator
    {
        byte[] GenerateReport(WeldingReportRequest request, Dictionary<string, List<string>> photoMap);
    }

    public class ExcelReportGenerator : IExcelReportGenerator
    {
        private const string TemplatePath = "Resources/Templates/WeldingReportTemplate.xlsx";
        private const string WorksheetName = "Отчет";
        private const int MaxRowHeight = 200; // Максимальная высота строки
        private readonly ILogger<ExcelReportGenerator> _logger;

        public ExcelReportGenerator(ILogger<ExcelReportGenerator> logger)
        {
            _logger = logger;
        }

        public byte[] GenerateReport(WeldingReportRequest request, Dictionary<string, List<string>> photoMap)
        {
            var templateFile = new FileInfo(TemplatePath);

            if (!templateFile.Exists)
                throw new FileNotFoundException("Шаблон отчета не найден", TemplatePath);

            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[WorksheetName];
            FillData(worksheet, request, photoMap);
            return package.GetAsByteArray();
        }

        private void FillData(ExcelWorksheet worksheet, WeldingReportRequest request, Dictionary<string, List<string>> photoMap)
        {
            const int startRow = 3;
            const int photoColumn = 9;
            _logger.LogInformation($"Total joints: {request.Joints.Count}");

            for (var i = 0; i < request.Joints.Count; i++)
            {
                var joint = request.Joints[i];
                var row = startRow + i;

                worksheet.Cells[row, 1].Value = request.ReportNumber;
                worksheet.Cells[row, 2].Value = joint.EquipmentType;
                worksheet.Cells[row, 3].Value = joint.PipelineNumber;
                worksheet.Cells[row, 4].Value = joint.CompanyName;
                worksheet.Cells[row, 5].Value = joint.JointNumber;
                worksheet.Cells[row, 6].Value = joint.DiameterMm;
                worksheet.Cells[row, 7].Value = joint.LengthMeters;
                worksheet.Cells[row, 8].Value = joint.DiameterMm / 25.4;

                // Добавление всех фото
                double maxHeight = 50; // Высота будет хотя бы 50
                double currentWidth = 0; // Отслеживает ширину занятого пространства

                if (!string.IsNullOrEmpty(joint.JointNumber) && photoMap.TryGetValue(joint.JointNumber, out var photos))
                {
                    for (int j = 0; j < photos.Count; j++)
                    {
                        var imgSize = InsertImage(worksheet, row, photoColumn, photos[j], currentWidth);
                        currentWidth += imgSize.width + 5; // Добавляем небольшой отступ
                        maxHeight = Math.Max(maxHeight, imgSize.height);
                    }
                }

                worksheet.Row(row).Height = Math.Min(MaxRowHeight, maxHeight);
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

            double scale = Math.Min(1.0, MaxRowHeight / (double)bitmap.Height);
            int newWidth = (int)(bitmap.Width * scale);
            int newHeight = (int)(bitmap.Height * scale);

            var picture = worksheet.Drawings.AddPicture($"Photo_{row}_{column}_{xOffset}", imagePath);
            picture.SetPosition(row - 1, 0, column - 1, (int)xOffset);
            picture.SetSize(newWidth, newHeight);

            return (newWidth, newHeight);
        }
    }
}
