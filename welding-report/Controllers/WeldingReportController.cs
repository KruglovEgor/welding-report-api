using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Swashbuckle.AspNetCore.Annotations;
using welding_report.Models;
using welding_report.Services;
using welding_report.Services.Request;
using welding_report.Services.Welding;

[ApiController]
[Route("api/[controller]")]
public class WeldingReportController : ControllerBase
{
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<WeldingReportController> _logger;
    //private readonly IWeldingExcelReportGenerator _weldingExcelGenerator;
    private readonly IEmailServiceFactory _emailServiceFactory;
    private readonly AppSettings _appSettings;
    private readonly IRedmineServiceFactory _redmineServiceFactory;
    private readonly INumberToText _numberToText;
    private readonly IRequestWordReportGenerator _wordReportGenerator;

    public WeldingReportController(
        IWebHostEnvironment env,
        ILogger<WeldingReportController> logger,
        IEmailServiceFactory emailServiceFactory,
        IOptions<AppSettings> appSettings,
        IRedmineServiceFactory redmineServiceFactory,
        INumberToText numberToText,
        IRequestWordReportGenerator wordReportGenerator
        )
    {
        _env = env;
        _logger = logger;
        _emailServiceFactory = emailServiceFactory;
        _appSettings = appSettings.Value;
        _redmineServiceFactory = redmineServiceFactory;
        _numberToText = numberToText;
        _wordReportGenerator = wordReportGenerator;
    }

    [HttpGet("generate-issue-from-request")]
    public async Task<IActionResult> GenerateIssueFromRequest(
        [FromQuery] int issueId = 45,
        [FromQuery] string apiKey = "secret"
    )
    {
        string reportPath = string.Empty;
        try
        {
            _logger.LogInformation("Получен запрос на генерацию документа из заявки. IssueId: {IssueId}", issueId);

            _logger.LogInformation("Начало обработки запроса для заявки {IssueId}", issueId);
            var requestRedmineService = _redmineServiceFactory.CreateRequestService(apiKey);
            var reportData = await requestRedmineService.GetRequestReportDataAsync(issueId);
            _numberToText.FillCostText(reportData);

            if (string.IsNullOrEmpty(reportData.CuratorEmail))
            {
                _logger.LogWarning("CuratorEmail отсутствует для заявки {IssueId}", issueId);
            }

            _logger.LogInformation("Начало создания документа Word для заявки {IssueId}", issueId);
            // Генерация документа
            var docBytes = _wordReportGenerator.GenerateRequestReport(reportData);
            _logger.LogInformation("Документ Word успешно сгенерирован для заявки {IssueId}", issueId);

            // Сохранение файла
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportData.Name}.docx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, docBytes);

            // Возвращаем файл и запускаем задачу на удаление после отправки
            var result = File(
                docBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"{reportData.Name}.docx"
            );

            // Запускаем удаление файла после возврата результата
            _ = Task.Run(async () => {
                // Небольшая задержка для завершения отправки файла
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });

            _logger.LogInformation("Документ успешно сгенерирован и отправлен клиенту для заявки {IssueId}", issueId);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");

            // Удаляем файл в случае ошибки, если он был создан
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }


    [HttpGet("generate-issue-from-welding")]
    public async Task<IActionResult> GenerateIssueFromWelding(
        [FromQuery] int issueId = 6,
        [FromQuery] int projectIdentifier = 1,
        [FromQuery] string apiKey = "secret",
        [FromQuery] bool sendMail = false)
    {
        string reportPath = string.Empty;
        try
        {
            _logger.LogInformation("Получен запрос на генерацию отчета по сварке. IssueId: {IssueId}, ProjectIdentifier: {ProjectIdentifier}, SendMail: {SendMail}", issueId, projectIdentifier, sendMail);
            var weldingRedmineService = _redmineServiceFactory.CreateWeldingService(apiKey);
            var weldingExcelGenerator = _redmineServiceFactory.CreateWeldingExcelGenerator(apiKey);
            _logger.LogInformation("Начало обработки запроса для отчета по сварке {IssueId}", issueId);
            var reportData = await weldingRedmineService.GetWeldingIssueDataAsync(projectIdentifier, issueId);
            _logger.LogInformation("Начало создания Excel документа для отчета по сварке {IssueId}", issueId);
            var excelBytes = await weldingExcelGenerator.GenerateIssueReport(reportData);
            _logger.LogInformation("Excel документ успешно сгенерирован для отчета по сварке {IssueId}", issueId);


            // Сохранение отчета
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{reportData.ReportNumber}.xlsx"
            );
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                var weldingEmailService = _emailServiceFactory.CreateWeldingEmailService(apiKey);
                _logger.LogInformation("Начало отправки отчета {IssueId} по сварке по email", issueId);
                await weldingEmailService.SendRedmineReportAsync(excelBytes, reportData.ReportNumber, apiKey);
                _logger.LogInformation("Отчет по сварке {IssueId} успешно отправлен по email", issueId);

                // Удаление файла после отправки email
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }

                return Ok("Отчет успешно создан и отправлен по электронной почте");
            }

            // Возвращаем файл и запускаем задачу на удаление после отправки
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{reportData.ReportNumber}.xlsx"
            );

            // Запускаем удаление файла после возврата результата
            _ = Task.Run(async () => {
                // Небольшая задержка для завершения отправки файла
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });
            _logger.LogInformation("Документ успешно сгенерирован и отправлен клиенту для отчета по сварке {IssueId}", issueId);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");

            // Удаляем файл в случае ошибки, если он был создан
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }

    [HttpGet("generate-project-from-welding")]
    public async Task<IActionResult> GenerateProjectFromWelding(
        [FromQuery] int projectIdentifier = 1,
        [FromQuery] string apiKey = "secret",
        [FromQuery] bool sendMail = false)
    {
        string reportPath = string.Empty;
        try
        {
            _logger.LogInformation("Получен запрос на генерацию отчета по проекту сварки. ProjectIdentifier: {ProjectIdentifier}, SendMail: {SendMail}", projectIdentifier, sendMail);
            var weldingRedmineService = _redmineServiceFactory.CreateWeldingService(apiKey);
            var weldingExcelGenerator = _redmineServiceFactory.CreateWeldingExcelGenerator(apiKey);
            _logger.LogInformation("Начало обработки запроса для отчета по проекту {ProjectIdentifier}", projectIdentifier);
            var projectData = await weldingRedmineService.GetProjectReportDataAsync(projectIdentifier);
            _logger.LogInformation("Начало создания Excel документа для проекта {ProjectName}", projectData.Name);
            var excelBytes = await weldingExcelGenerator.GenerateProjectReport(projectData);
            _logger.LogInformation("Excel документ успешно сгенерирован для проекта {ProjectName}", projectData.Name);

            // Сохранение отчета
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                $"{projectData.Name}.xlsx"
            );

            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);

            if (sendMail)
            {
                var weldingEmailService = _emailServiceFactory.CreateWeldingEmailService(apiKey);
                _logger.LogInformation("Начало отправки отчета по проекту {ProjectName} по email", projectData.Name);
                await weldingEmailService.SendRedmineReportAsync(excelBytes, projectData.Name, apiKey);
                _logger.LogInformation("Отчет по проекту {ProjectName} успешно отправлен по email", projectData.Name);

                // Удаление файла после отправки email
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }

                return Ok("Отчет успешно создан и отправлен по электронной почте");
            }

            // Возвращаем файл и запускаем задачу на удаление после отправки
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{projectData.Name}.xlsx"
            );

            // Запускаем удаление файла после возврата результата
            _ = Task.Run(async () => {
                // Небольшая задержка для завершения отправки файла
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });
            _logger.LogInformation("Документ успешно сгенерирован и отправлен клиенту для проекта {ProjectName}", projectData.Name);

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");

            // Удаляем файл в случае ошибки, если он был создан
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }


    [HttpGet("generate-group-from-supr")]
    public async Task<IActionResult> GenerateGroupFromSupr(
        [FromQuery] string projectIdentifier = "test_project",
        [FromQuery] string apiKey = "secret",
        [FromQuery] int applicationNumber = 222
        )
    {
        string reportPath = string.Empty;
        try
        {
            _logger.LogInformation("Получен запрос на генерацию отчета SUPR. ProjectIdentifier: {ProjectIdentifier}, ApplicationNumber: {ApplicationNumber}", projectIdentifier, applicationNumber);
            var suprRedmineService = _redmineServiceFactory.CreateSuprService(apiKey);
            _logger.LogInformation("Начало обработки запроса SUPR для проекта {ProjectIdentifier}, заявка {ApplicationNumber}", projectIdentifier, applicationNumber);
            var projectData = await suprRedmineService.GetSuprGroupReportDataAsync(projectIdentifier, applicationNumber);
            var excelGenerator = _redmineServiceFactory.CreateSuprExcelGenerator();
            _logger.LogInformation("Начало создания Excel документа SUPR для предприятия {Factory}, заявка {ApplicationNumber}", projectData.Factory, applicationNumber);
            var excelBytes = await excelGenerator.GenerateGroupReport(projectData);
            _logger.LogInformation("Excel документ SUPR успешно сгенерирован для предприятия {Factory}, заявка {ApplicationNumber}", projectData.Factory, applicationNumber);


            // Определение пути для временного сохранения файла
            string fileName = $"SUPR_{projectData.Factory}_{applicationNumber}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            reportPath = Path.Combine(
                _env.ContentRootPath,
                _appSettings.ReportStoragePath,
                fileName
            );

            // Создаем директорию, если не существует
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            await System.IO.File.WriteAllBytesAsync(reportPath, excelBytes);


            // Возвращаем файл для скачивания
            var result = File(
                excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName
            );

            // Удаляем временный файл после отправки
            _ = Task.Run(async () => {
                await Task.Delay(1000);
                if (System.IO.File.Exists(reportPath))
                {
                    System.IO.File.Delete(reportPath);
                }
            });
            _logger.LogInformation("Документ SUPR успешно сгенерирован и отправлен клиенту для предприятия {Factory}, заявка {ApplicationNumber}", projectData.Factory, applicationNumber);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Ошибка генерации отчета");

            // Удаляем файл в случае ошибки, если он был создан
            if (!string.IsNullOrEmpty(reportPath) && System.IO.File.Exists(reportPath))
            {
                System.IO.File.Delete(reportPath);
            }

            return StatusCode(500, ex.Message);
        }
    }
}