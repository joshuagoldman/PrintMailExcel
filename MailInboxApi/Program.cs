using MailInboxApi.Data.Repositories;
using MailInboxApi.Services;

var builder = WebApplication.CreateBuilder(args);

string env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");

string appJsonSettingsFile = 
    String.IsNullOrEmpty(env) ? "appsettings.json" : $"appsettings.{env}.json";

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddSingleton<IInboxExcelRepository, InboxExcelRepository>();
builder.Services.AddSingleton<IGetAllInboxExcelService, GetAllInboxExcelService>();

IConfigurationRoot Configuration = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile(appJsonSettingsFile, optional: true,reloadOnChange: true)
                        .AddEnvironmentVariables()
                        .Build();

builder.Services.AddSingleton<IConfiguration>(Configuration);

var app = builder.Build();
// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
