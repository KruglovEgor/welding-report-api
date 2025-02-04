using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.FileProviders;
using OfficeOpenXml;


var builder = WebApplication.CreateBuilder(args);

// Настройка путей для загрузок
var uploadsPath = Path.Combine(builder.Environment.ContentRootPath, "uploads");
if (!Directory.Exists(uploadsPath))
{
    Directory.CreateDirectory(uploadsPath);
    Console.WriteLine($"Created uploads directory at: {uploadsPath}");
}

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
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