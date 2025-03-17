using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Options;
using welding_report.Models;

namespace welding_report.Services
{

    public interface IRequestWordReportGenerator
    {
        byte[] GenerateRequestReport(RequestReportData data);
    }

    public class RequestWordReportGenerator : IRequestWordReportGenerator
    {

        private readonly string _templatePath;
        private readonly ILogger<RequestWordReportGenerator> _logger;


        public RequestWordReportGenerator(
           ILogger<RequestWordReportGenerator> logger,
           IOptions<AppSettings> appSettings
            )
        {
            _logger = logger;
            _templatePath = System.IO.Path.Combine(
                appSettings.Value.TemplatePath,
                "RequestReportTemplate.docx"
            );
        }
        public byte[] GenerateRequestReport(RequestReportData data)
        {
            try
            {
                var templateBytes = File.ReadAllBytes(_templatePath);
                using var stream = new MemoryStream();
                stream.Write(templateBytes, 0, templateBytes.Length);

                using (var doc = WordprocessingDocument.Open(stream, true))
                {
                    var mainPart = doc.MainDocumentPart;
                    ReplaceContentControls(mainPart, data);
                }

                return stream.ToArray();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Ошибка генерации отчета");
                throw;
            }
        }

        private void ReplaceContentControls(MainDocumentPart mainPart, RequestReportData data)
        {
            foreach (var sdt in mainPart.Document.Descendants<SdtElement>())
            {
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (string.IsNullOrEmpty(tag)) continue;

                var text = sdt.Descendants<Text>().FirstOrDefault();
                if (text == null) continue;

                switch (tag)
                {
                    case "Name":
                        text.Text = data.Name;
                        break;
                    case "CustomerName":
                        text.Text = data.CustomerName;
                        break;
                    case "CustomerEmail":
                        text.Text = data.CustomerEmail;
                        break;

                    case "Theme":
                        text.Text = data.Theme;
                        break;

                    case "Aim":
                        text.Text = data.Aim;
                        break;

                    case "RequestDate":
                        text.Text = data.RequestDate;
                        break;
                    case "CuratorName":
                        text.Text = data.CuratorName;
                        break;
                    case "CuratorEmail":
                        text.Text = data.CuratorEmail;
                        break;
                    case "PlanStartDateText":
                        text.Text = data.PlanStartDateText;
                        break;

                    case "PlanEndDateText":
                        text.Text = data.PlanEndDateText;
                        break;

                    case "Cost":
                        text.Text = data.Cost;
                        break;

                    case "CostText":
                        text.Text = data.CostText;
                        break;

                    case "OwnCost":
                        text.Text = data.OwnCost;
                        break;

                    case "SubCost":
                        text.Text = data.SubCost;
                        break;

                    case "MaterialCost":
                        text.Text = data.MaterialCost;
                        break;

                    case "OtherCost":
                        text.Text = data.OtherCost;
                        break;

                    default:
                        _logger.LogWarning("Неизвестный тег элемента управления: {Tag}", tag);
                        break;
                }
            }
        }
    }
}
