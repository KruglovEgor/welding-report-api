using OfficeOpenXml;
using OfficeOpenXml.Style;
using Microsoft.Extensions.Options;
using System.Drawing;
using welding_report.Models;
using DocumentFormat.OpenXml.Spreadsheet;

namespace welding_report.Services.Supr
{
    public interface ISuprExcelReportGenerator
    {
        Task<byte[]> GenerateGroupReport(SuprGroupReportData data);
    }

    public class SuprExcelReportGenerator : ISuprExcelReportGenerator
    {
        private readonly string _templatePath;
        private readonly ILogger<SuprExcelReportGenerator> _logger;

        public SuprExcelReportGenerator(
            ILogger<SuprExcelReportGenerator> logger,
            IOptions<AppSettings> appSettings)
        {
            _logger = logger;
            _templatePath = Path.Combine(
                appSettings.Value.TemplatePath,
                "SuprReportTemplate.xlsx");
        }

        public async Task<byte[]> GenerateGroupReport(SuprGroupReportData data)
        {
            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[0]; // Берем первый лист из шаблона

            int count = data.suprIssueReportDatas.Count;
            int startRow = 10;


            // Заполняем шапку отчета
            worksheet.Cells[startRow, 3].Value = data.Factory;
            worksheet.Cells[startRow, 4].Value = data.InstallationName;
            worksheet.Cells[startRow, 5].Value = data.TechPositionName;
            worksheet.Cells[startRow, 6].Value = data.EquipmentUnitNumber;
            worksheet.Cells[startRow, 8].Value = data.MarkAndManufacturer;

            var groupParametersColumns = new[] { 3, 4, 5, 6, 7, 8 };

            foreach(var i in groupParametersColumns)
            {
                var rangeToMerge = worksheet.Cells[startRow, i, startRow+count-1, i];
                rangeToMerge.Merge = true;
                ApplyBorders(rangeToMerge);
            }

            int currentRow = startRow;

            // Заполняем данные
            foreach (var issue in data.suprIssueReportDatas.OrderBy(x => x.Key))
            {
                worksheet.Cells[currentRow, 2].Value = issue.Key;
                worksheet.Cells[currentRow, 9].Value = issue.Value.Detail;
                worksheet.Cells[currentRow, 10].Value = issue.Value.ScanningPeriod;
                worksheet.Cells[currentRow, 11].Value = issue.Value.Condition;
                worksheet.Cells[currentRow, 12].Value = issue.Value.Priority;
                worksheet.Cells[currentRow, 13].Value = issue.Value.JobType;


                ApplyBorders(worksheet.Cells[currentRow, 2]);
                ApplyBorders(worksheet.Cells[currentRow, 9, currentRow, 13]);

                worksheet.Row(currentRow).CustomHeight = false;
                currentRow++;
            }


            // Добавляем две константные строки ниже таблицы
            currentRow += 3; 

            worksheet.Cells[currentRow, 4].Value = "Исполнитель";
            worksheet.Cells[currentRow+1, 4].Value = "Генеральный директор ";
            worksheet.Cells[currentRow + 2, 4].Value = "ООО \"ЛИНК\"";
            worksheet.Cells[currentRow + 4, 4].Value = "________________/М.Р. Усманов";
            worksheet.Cells[currentRow, 4, currentRow + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[currentRow, 4, currentRow + 4, 4].Style.Font.Size = 16;
            worksheet.Cells[currentRow, 4, currentRow + 4, 4].Style.WrapText = false;

            worksheet.Cells[currentRow, 10].Value = "Заказчик";
            worksheet.Cells[currentRow + 1, 10].Value = "Генеральный директор ";
            worksheet.Cells[currentRow + 2, 10].Value = "ООО \"ЛУКОЙЛ-Нижегороднефтеоргсинтез\"";
            worksheet.Cells[currentRow + 4, 10].Value = "________________/С.М. Андронов";
            worksheet.Cells[currentRow, 10, currentRow + 4, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[currentRow, 10, currentRow + 4, 10].Style.Font.Size = 16;
            worksheet.Cells[currentRow, 10, currentRow + 4, 10].Style.WrapText = false;

            return package.GetAsByteArray();
        }


        // Вспомогательный метод для добавления границ
        private void ApplyBorders(ExcelRange range)
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }
    }
}
