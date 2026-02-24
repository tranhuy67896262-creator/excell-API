using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// ✅ Đặt license đúng cách cho EPPlus 8.4.x
ExcelPackage.LicenseContext = LicenseContext.Commercial;

// Nếu bạn chỉ dùng phi thương mại
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
// --- Cấu hình CORS ---
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowFrontend", policy =>
    {
        policy
            .WithOrigins("http://127.0.0.1:5500", "http://localhost:5500")
            .AllowAnyHeader()
            .AllowAnyMethod()
            .AllowCredentials();
    });
});

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

app.UseCors("AllowFrontend");

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();
