using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Connectors.AzureOpenAI;
using Microsoft.AspNetCore.Mvc;
using PhraseXross;
using PhraseXross.Dialogs;
using Microsoft.Bot.Builder.Dialogs;

var builder = WebApplication.CreateBuilder(args);

// Load .env manually (simple parser) before building services so Environment.GetEnvironmentVariable works.
var envFile = Path.Combine(AppContext.BaseDirectory, ".env");
if (File.Exists(envFile))
{
    foreach (var line in File.ReadAllLines(envFile))
    {
        var trimmed = line.Trim();
        if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#")) continue;
        var idx = trimmed.IndexOf('=');
        if (idx <= 0) continue;
        var key = trimmed[..idx].Trim();
        var value = trimmed[(idx + 1)..].Trim();
        // Only set if not already set in process env
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable(key)))
        {
            Environment.SetEnvironmentVariable(key, value);
        }
    }
}

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddControllers();
// Graph API HttpClient (basic)
builder.Services.AddHttpClient("graph", c =>
{
    c.Timeout = TimeSpan.FromSeconds(100);
});
builder.Services.AddSingleton<PhraseXross.Services.OneDriveExcelService>();
// ---- Adapter Pattern A (standard) ----
// Collect auth settings
var authSettings = new Dictionary<string, string?>();
void AddIf(string key)
{
    var val = Environment.GetEnvironmentVariable(key);
    if (!string.IsNullOrWhiteSpace(val)) authSettings[key] = val;
}
AddIf("MicrosoftAppId");
AddIf("MicrosoftAppPassword");
AddIf("MicrosoftAppTenantId");
AddIf("MicrosoftAppType");

builder.Services.AddSingleton<BotFrameworkAuthentication>(sp =>
{
    var logger = sp.GetRequiredService<ILoggerFactory>().CreateLogger("AuthConfig");
    logger.LogInformation("[DEBUG][BOOT] Using BotFrameworkAuthentication (Pattern A)");
    var cfg = new ConfigurationBuilder().AddInMemoryCollection(authSettings).Build();
    return new ConfigurationBotFrameworkAuthentication(cfg);
});

builder.Services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();
// Dialog / Bot registration
builder.Services.AddSingleton<MainDialog>();
builder.Services.AddTransient<IBot>(sp =>
{
    var kernel = sp.GetService<Kernel>();
    var userState = sp.GetRequiredService<UserState>();
    var oneDrive = sp.GetService<PhraseXross.Services.OneDriveExcelService>();
    var dialog = sp.GetRequiredService<MainDialog>();
    var convo = sp.GetService<ConversationState>();
    return new SimpleBot(kernel, userState, oneDrive, dialog, convo);
});

// Bot State for multi-turn conversation (yes/no confirmation flow)
builder.Services.AddSingleton<IStorage, MemoryStorage>();
builder.Services.AddSingleton<UserState>(sp => new UserState(sp.GetRequiredService<IStorage>()));
builder.Services.AddSingleton<ConversationState>(sp => new ConversationState(sp.GetRequiredService<IStorage>()));

// Semantic Kernel (optional)
// ENABLE_SK=true かつ AOAI_* が揃っている場合のみ Kernel を登録
bool IsSkEnabled(IConfiguration cfg)
{
    var flag = Environment.GetEnvironmentVariable("ENABLE_SK") ?? cfg["Features:ENABLE_SK"] ?? "false";
    Console.WriteLine($"[DEBUG] ENABLE_SK: {flag}");
    return string.Equals(flag, "true", StringComparison.OrdinalIgnoreCase);
}

if (IsSkEnabled(builder.Configuration))
{
    // 事前に構成が揃っているか確認し、不足時は登録スキップ
    string? endpoint = Environment.GetEnvironmentVariable("AOAI_ENDPOINT");
    string? apiKey = Environment.GetEnvironmentVariable("AOAI_API_KEY");
    string? deployment = Environment.GetEnvironmentVariable("AOAI_DEPLOYMENT");

    string Mask(string? secret)
    {
        if (string.IsNullOrEmpty(secret)) return "(null)";
        var trimmed = secret.Trim();
        if (trimmed.Length <= 4) return new string('*', trimmed.Length);
        // 先頭4文字 + *** + 末尾2文字 + (len=N) を表示し中身を漏らさない
        return $"{trimmed.Substring(0,4)}***{trimmed.Substring(trimmed.Length-2,2)}(len={trimmed.Length})";
    }

    Console.WriteLine($"[DEBUG] AOAI_ENDPOINT: {endpoint}");
    Console.WriteLine($"[DEBUG] AOAI_API_KEY: {Mask(apiKey)}");
    Console.WriteLine($"[DEBUG] AOAI_DEPLOYMENT: {deployment}");

    if (string.IsNullOrWhiteSpace(endpoint) || string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(deployment))
    {
        Console.WriteLine("[SK] ENABLE_SK=true ですが AOAI_ENDPOINT/AOAI_API_KEY/AOAI_DEPLOYMENT のいずれかが未設定のため、Semantic Kernel の登録をスキップします。");
    }
    else
    {
        builder.Services.AddSingleton<Kernel>(sp =>
        {
            var logger = sp.GetRequiredService<ILoggerFactory>().CreateLogger("SK");
            var kb = Kernel.CreateBuilder();
            kb.AddAzureOpenAIChatCompletion(deployment, endpoint, apiKey);
            var kernel = kb.Build();
            logger.LogInformation("[SK] Semantic Kernel を登録しました (deployment={Deployment}, endpoint={Endpoint})", deployment, endpoint);
            return kernel;
        });
    }
}

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

// Minimal API for triggering OneDrive Excel generation (manual test)
app.MapPost("/onedrive/excel/generate", async ([FromServices] PhraseXross.Services.OneDriveExcelService svc, CancellationToken ct) =>
{
    // 進捗コールバック不要のため null を指定（新シグネチャ対応）
    var result = await svc.CreateAndFillExcelAsync(null, null, ct);
    return result.IsSuccess ? Results.Ok(new { result.WebUrl, result.FileName }) : Results.BadRequest(new { error = result.Error });
})
.WithName("GenerateOneDriveExcel")
.WithOpenApi();

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
