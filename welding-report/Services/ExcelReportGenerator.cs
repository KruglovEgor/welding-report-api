namespace welding_report.Services
{
    using OfficeOpenXml;
    using System.Drawing;
    using welding_report.Models;

    public interface IExcelReportGenerator
    {
        byte[] GenerateReport(WeldingReportRequest request, Dictionary<string, List<string>> photoMap);
    }

    public class ExcelReportGenerator : IExcelReportGenerator
    {
        private const string TemplatePath = "Resources/Templates/WeldingReportTemplate.xlsx";
        private const string WorksheetName = "Отчет";
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

                if (!string.IsNullOrEmpty(joint.JointNumber) &&
                    photoMap.TryGetValue(joint.JointNumber, out var photos) &&
                    photos.Count > 0)
                {
                    InsertImage(worksheet, row, photos.First());
                }
                else
                {
                    _logger.LogWarning($"No photos found for joint {joint.JointNumber}");
                }
            }
        }

        private void InsertImage(ExcelWorksheet worksheet, int row, string imagePath)
        {
            try
            {
                using var imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                var picture = worksheet.Drawings.AddPicture($"Photo{row}", imageStream);
                picture.SetPosition(row - 1, 0, 9, 0);
                picture.SetSize(100, 100);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to insert image {imagePath}");
            }
        }

    }
}
