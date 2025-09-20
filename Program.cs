using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.DependencyInjection;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddControllers();
builder.Services.AddSingleton<CloudAdapter, CloudAdapter>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<CloudAdapter>>();

    // Retrieve credentials from environment variables
    // Log environment variables for debugging purposes with detailed information
    var appId = Environment.GetEnvironmentVariable("MicrosoftAppId");
    var appPassword = Environment.GetEnvironmentVariable("MicrosoftAppPassword");
    var tenantId = Environment.GetEnvironmentVariable("MicrosoftAppTenantId"); // 統一してMicrosoftAppTenantIdを使用

    var appType = Environment.GetEnvironmentVariable("MicrosoftAppType") ?? "SingleTenant"; // 追加で取得

    // Log environment variables for debugging purposes
    logger.LogInformation("[DEBUG] MicrosoftAppId: {AppId}", appId);
    logger.LogInformation("[DEBUG] MicrosoftAppPassword: {AppPassword}", appPassword);
    logger.LogInformation("[DEBUG] MicrosoftAppTenantId: {AppTenantId}", tenantId);
    logger.LogInformation("[DEBUG] MicrosoftAppType: {AppType}", appType);

    // Log detailed authentication request information
    logger.LogInformation("[DEBUG] Preparing authentication request with the following details:");
    logger.LogInformation("[DEBUG] Authentication Endpoint: https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token", tenantId);
    logger.LogInformation("[DEBUG] MicrosoftAppId: {AppId}", appId);
    logger.LogInformation("[DEBUG] MicrosoftTenantId: {TenantId}", tenantId);
    logger.LogInformation("[DEBUG] MicrosoftAppPassword: {AppPassword}", appPassword);
    logger.LogInformation("[DEBUG] MicrosoftAppType: {AppType}", Environment.GetEnvironmentVariable("MicrosoftAppType"));
    logger.LogInformation("[DEBUG] MicrosoftAppTenantId: {AppTenantId}", Environment.GetEnvironmentVariable("MicrosoftAppTenantId"));

    var botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(new ConfigurationBuilder()
        .AddInMemoryCollection(new Dictionary<string, string?>
        {
            { "MicrosoftAppId", appId },
            { "MicrosoftAppPassword", appPassword },
            { "MicrosoftAppType", appType },
            { "MicrosoftAppTenantId", tenantId }
        })
        .Build());

    return new CloudAdapter(botFrameworkAuthentication, logger);
});
builder.Services.AddTransient<IBot, SimpleBot>(); // Replace SimpleBot with your bot implementation

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

var summaries = new[]
{
    "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
};

app.MapGet("/weatherforecast", () =>
{
    var forecast =  Enumerable.Range(1, 5).Select(index =>
        new WeatherForecast
        (
            DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
            Random.Shared.Next(-20, 55),
            summaries[Random.Shared.Next(summaries.Length)]
        ))
        .ToArray();
    return forecast;
})
.WithName("GetWeatherForecast")
.WithOpenApi();

app.Run();

record WeatherForecast(DateOnly Date, int TemperatureC, string? Summary)
{
    public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);
}
