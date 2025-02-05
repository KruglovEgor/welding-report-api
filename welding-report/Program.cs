using Microsoft.Extensions.FileProviders;
using Microsoft.OpenApi.Models;
using OfficeOpenXml;
using welding_report.Services;


var builder = WebApplication.CreateBuilder(args);

// ��������� ����� ��� ��������
var uploadsPath = Path.Combine(builder.Environment.ContentRootPath, "uploads");
if (!Directory.Exists(uploadsPath))
{
    Directory.CreateDirectory(uploadsPath);
    Console.WriteLine($"Created uploads directory at: {uploadsPath}");
}

builder.Services.AddControllers();
builder.Services.AddScoped<IExcelReportGenerator, ExcelReportGenerator>();
builder.Services.AddSingleton<ExcelReportGenerator>();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo { Title = "Welding Report API", Version = "v1" });

    // ��������� ��������� multipart/form-data
    c.OperationFilter<welding_report.Models.FormDataOperationFilter>();

    // �������� ���������
    c.EnableAnnotations();
}); 


// ���������� ������������ ��������
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ��� ��������������� �������������s

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