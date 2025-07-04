using OfficeOpenXml;
using OfficeOpenXml.Style;
using Microsoft.Extensions.Options;
using System.Drawing;
using welding_report.Models;
using DocumentFormat.OpenXml.Spreadsheet;
using welding_report.Models.Supr;
using System.Security.Cryptography.Xml;

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
        private readonly SuprSignatures _signatures;

        public SuprExcelReportGenerator(
            ILogger<SuprExcelReportGenerator> logger,
            IOptions<AppSettings> appSettings,
            IOptions<SuprSignatures> signatures)
        {
            _logger = logger;
            _templatePath = Path.Combine(
                appSettings.Value.TemplatePath,
                "SuprReportTemplate.xlsx");
            _signatures = signatures.Value;
        }

        public async Task<byte[]> GenerateGroupReport(SuprGroupReportData data)
        {
            var templateFile = new FileInfo(_templatePath);
            using var package = new ExcelPackage(templateFile);
            var worksheet = package.Workbook.Worksheets[0]; // Берем первый лист из шаблона

            // Заполняем номер заявки
            worksheet.Cells[5, 8].Value += $"{data.ApplicationNumber} от __.__ 20__г";
            worksheet.Cells[2, 12].Value += data.ContractNumber + ".";

            int count = data.suprIssueReportDatas.Count;
            int startRow = 10;

            int currentRow = startRow;

            // Заполняем данные и запоминаем значения для каждой строки
            var cellValues = new Dictionary<int, Dictionary<int, string>>();
            worksheet.Cells[7, 2, 7, 13].AutoFilter = true;

            // Заполняем данные
            foreach (var issue in data.suprIssueReportDatas
                .OrderBy(x => x.Value.InstallationName)
                .ThenBy(x => x.Value.TechPositionName))
            {
                worksheet.Cells[currentRow, 2].Value = issue.Key;
                worksheet.Cells[currentRow, 3].Value = data.Factory;
                worksheet.Cells[currentRow, 9].Value = issue.Value.Detail;
                worksheet.Cells[currentRow, 10].Value = issue.Value.ScanningPeriod;
                worksheet.Cells[currentRow, 11].Value = issue.Value.Condition;
                worksheet.Cells[currentRow, 12].Value = issue.Value.Priority;
                worksheet.Cells[currentRow, 13].Value = issue.Value.JobType;

                worksheet.Cells[currentRow, 4].Value = issue.Value.InstallationName;
                worksheet.Cells[currentRow, 5].Value = issue.Value.TechPositionName;
                worksheet.Cells[currentRow, 6].Value = issue.Value.EquipmentUnitNumber;
                worksheet.Cells[currentRow, 7].Value = issue.Value.EquipmentType;
                worksheet.Cells[currentRow, 8].Value = issue.Value.MarkAndManufacturer;



                ApplyBorders(worksheet.Cells[currentRow, 2, currentRow, 13]);
                //ApplyBorders(worksheet.Cells[currentRow, 4, currentRow, 8]);
                //ApplyBorders(worksheet.Cells[currentRow, 9, currentRow, 13]);

                worksheet.Row(currentRow).CustomHeight = false;
                currentRow++;
            }

            // Добавляем две константные строки ниже таблицы
            currentRow += 3; 

            worksheet.Cells[currentRow, 4].Value = _signatures.Executor["Role"];
            worksheet.Cells[currentRow+1, 4].Value = _signatures.Executor["JobTitle"];
            worksheet.Cells[currentRow + 2, 4].Value = _signatures.Executor["Company"];
            worksheet.Cells[currentRow + 4, 4].Value = _signatures.Executor["Signature"];
            worksheet.Cells[currentRow, 4, currentRow + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[currentRow, 4, currentRow + 4, 4].Style.Font.Size = 16;
            worksheet.Cells[currentRow, 4, currentRow + 4, 4].Style.WrapText = false;

            worksheet.Cells[currentRow, 10].Value = _signatures.Customer["Role"];
            worksheet.Cells[currentRow + 1, 10].Value = _signatures.Customer["JobTitle"];
            worksheet.Cells[currentRow + 2, 10].Value = _signatures.Customer["Company"];
            worksheet.Cells[currentRow + 4, 10].Value = _signatures.Customer["Signature"];
            worksheet.Cells[currentRow, 10, currentRow + 4, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[currentRow, 10, currentRow + 4, 10].Style.Font.Size = 16;
            worksheet.Cells[currentRow, 10, currentRow + 4, 10].Style.WrapText = false;


            if (!string.IsNullOrEmpty(data.CustomerRepresentative))
            {
                worksheet.Cells[currentRow + 4, 10].Value = $"________________/{data.CustomerRepresentative}";
                //_logger.LogInformation("Customer representative: {Representative}", data.CustomerRepresentative);
            }
            if (!string.IsNullOrEmpty(data.CustomerCompany))
            {
                worksheet.Cells[currentRow + 2, 10].Value = data.CustomerCompany;
                //_logger.LogInformation("Customer company: {Company}", data.CustomerCompany);
            }


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
