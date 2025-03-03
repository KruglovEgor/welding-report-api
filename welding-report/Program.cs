using Microsoft.Extensions.FileProviders;
using Microsoft.OpenApi.Models;
using OfficeOpenXml;
using welding_report.Models;
using welding_report.Services;


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
builder.Services.AddScoped<IEmailService, EmailService>();
builder.Services.AddScoped<IExcelReportGenerator, ExcelReportGenerator>();
builder.Services.AddEndpointsApiExplorer();

builder.Configuration.AddUserSecrets<Program>();
builder.Services.Configure<EmailSettings>(builder.Configuration.GetSection("EmailSettings"));
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));
builder.Services.Configure<RedmineSettings>(builder.Configuration.GetSection("RedmineSettings"));
builder.Services.AddHttpClient<IRedmineService, RedmineService>();


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