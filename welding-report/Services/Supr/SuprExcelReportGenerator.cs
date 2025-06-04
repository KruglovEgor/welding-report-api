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

            // Заполняем номер заявки
            worksheet.Cells[5, 8].Value += $"{data.ApplicationNumber} от __.__ 20__г";


            int count = data.suprIssueReportDatas.Count;
            int startRow = 10;

            // Указываем завод
            worksheet.Cells[startRow, 3].Value = data.Factory;

            //TODO: убрать 7 как появится информация что делать с данной колонкой
            var groupParametersColumns = new[] { 3, 7 };

            foreach(var i in groupParametersColumns)
            {
                var rangeToMerge = worksheet.Cells[startRow, i, startRow+count-1, i];
                rangeToMerge.Merge = true;
                ApplyBorders(rangeToMerge);
            }

            int currentRow = startRow;

            // Заполняем данные и запоминаем значения для каждой строки
            var cellValues = new Dictionary<int, Dictionary<int, string>>();
            var columnsToMerge = new[] { 4, 5, 6, 8 }; // Колонки, для которых нужно объединять

            // Заполняем данные
            foreach (var issue in data.suprIssueReportDatas.OrderBy(x => x.Key))
            {
                worksheet.Cells[currentRow, 2].Value = issue.Key;
                worksheet.Cells[currentRow, 9].Value = issue.Value.Detail;
                worksheet.Cells[currentRow, 10].Value = issue.Value.ScanningPeriod;
                worksheet.Cells[currentRow, 11].Value = issue.Value.Condition;
                worksheet.Cells[currentRow, 12].Value = issue.Value.Priority;
                worksheet.Cells[currentRow, 13].Value = issue.Value.JobType;

                worksheet.Cells[currentRow, 4].Value = issue.Value.InstallationName;
                worksheet.Cells[currentRow, 5].Value = issue.Value.TechPositionName;
                worksheet.Cells[currentRow, 6].Value = issue.Value.EquipmentUnitNumber;
                worksheet.Cells[currentRow, 8].Value = issue.Value.MarkAndManufacturer;

                // Запоминаем значения для этой строки
                cellValues[currentRow] = new Dictionary<int, string>();
                foreach (var col in columnsToMerge)
                {
                    cellValues[currentRow][col] = worksheet.Cells[currentRow, col].Text;
                }

                ApplyBorders(worksheet.Cells[currentRow, 2]);
                ApplyBorders(worksheet.Cells[currentRow, 4, currentRow, 8]);
                ApplyBorders(worksheet.Cells[currentRow, 9, currentRow, 13]);

                worksheet.Row(currentRow).CustomHeight = false;
                currentRow++;
            }

            // Объединяем ячейки с одинаковыми значениями
            foreach (var col in columnsToMerge)
            {
                MergeConsecutiveCellsWithSameValue(worksheet, cellValues, startRow, currentRow - 1, col);
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


        // Метод для объединения последовательных ячеек с одинаковым значением
        private void MergeConsecutiveCellsWithSameValue(ExcelWorksheet worksheet, Dictionary<int, Dictionary<int, string>> cellValues, int startRow, int endRow, int column)
        {
            if (startRow >= endRow)
                return;

            int mergeStartRow = startRow;
            string currentValue = cellValues[mergeStartRow][column];

            for (int row = startRow + 1; row <= endRow + 1; row++) // +1 для обработки последней группы
            {
                bool isLastRow = row > endRow;
                bool valueChanged = isLastRow || currentValue != cellValues[row][column];

                if (valueChanged)
                {
                    // Если было две или больше последовательных ячеек с одинаковым значением, объединяем их
                    if (row - 1 > mergeStartRow)
                    {
                        var rangeToMerge = worksheet.Cells[mergeStartRow, column, row - 1, column];
                        rangeToMerge.Merge = true;
                        ApplyBorders(rangeToMerge);
                    }

                    // Начинаем новую группу ячеек
                    if (!isLastRow)
                    {
                        mergeStartRow = row;
                        currentValue = cellValues[row][column];
                    }
                }
            }
        }

    }
}
