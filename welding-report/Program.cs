using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Microsoft.Extensions.FileProviders;
using Microsoft.OpenApi.Models;
using OfficeOpenXml;
using welding_report.Models;
using welding_report.Models.Supr;
using welding_report.Services;
using welding_report.Services.Request;
using welding_report.Services.Welding;


//TODO add more comments
//TODO clean code
//TODO try to make Swagger look better and feel more comfortable for testing
//TODO make README file
//TODO look if we really need to save photos
//TODO fix cleanPhotos for local starting
//TODO maybe store some constant values in config file

var builder = WebApplication.CreateBuilder(args);

// Конфигурация путей из AppSettings
var appSettings = builder.Configuration.GetSection("AppSettings").Get<AppSettings>();
var uploadsPath = Path.Combine(builder.Environment.ContentRootPath, appSettings.UploadsFolder);
var reportsPath = Path.Combine(builder.Environment.ContentRootPath, appSettings.ReportStoragePath);

// Создание необходимых директорий
Directory.CreateDirectory(uploadsPath);
Directory.CreateDirectory(reportsPath);
Console.WriteLine($"Created directories:\nUploads: {uploadsPath}\nReports: {reportsPath}");


builder.Services.AddControllers();
//builder.Services.AddScoped<IEmailService, EmailService>();
//builder.Services.AddScoped<IWeldingExcelReportGenerator, WeldingExcelReportGenerator>();
builder.Services.AddScoped<INumberToText, NumberToText>();
builder.Services.AddScoped<IRequestWordReportGenerator, RequestWordReportGenerator>();
builder.Services.AddEndpointsApiExplorer();

builder.Configuration.AddUserSecrets<Program>();
builder.Services.Configure<EmailSettings>(builder.Configuration.GetSection("EmailSettings"));
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));
builder.Services.Configure<RedmineSettings>(builder.Configuration.GetSection("RedmineSettings"));
//builder.Services.AddHttpClient<IRedmineService, RedmineService>();
builder.Services.Configure<SuprSignatures>(builder.Configuration.GetSection("SuprSignatures"));


// Регистрация HttpClient'ов
builder.Services.AddHttpClient("Request")
    .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler
    {
        ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
    });
builder.Services.AddHttpClient("Welding")
    .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler
    {
        ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
    });
builder.Services.AddHttpClient("Supr")
    .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler
    {
        ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true
    });

// Регистрация фабрик сервисов
builder.Services.AddScoped<IRedmineServiceFactory, RedmineServiceFactory>();
builder.Services.AddScoped<IEmailServiceFactory, EmailServiceFactory>();


builder.Services.AddSwaggerGen();

// Установите лицензионный контекст
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Для некоммерческого использованияs

var app = builder.Build();
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(uploadsPath),
    RequestPath = "/uploads"
});

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();