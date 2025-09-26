using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Connectors.AzureOpenAI;
using Microsoft.AspNetCore.Mvc;

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
builder.Services.AddSingleton<CloudAdapter, CloudAdapter>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<CloudAdapter>>();

    // Retrieve credentials from environment variables
    // Log environment variables for debugging purposes with detailed information
    var appId = Environment.GetEnvironmentVariable("MicrosoftAppId");
    var appPassword = Environment.GetEnvironmentVariable("MicrosoftAppPassword");
    var tenantId = Environment.GetEnvironmentVariable("MicrosoftAppTenantId"); // 統一してMicrosoftAppTenantIdを使用

    var appType = Environment.GetEnvironmentVariable("MicrosoftAppType");

    // 判定: 資格情報が揃っているか
    var credsProvided = !string.IsNullOrWhiteSpace(appId) && !string.IsNullOrWhiteSpace(appPassword);

    // Log environment variables for debugging purposes (パスワードはマスク)
    //logger.LogInformation("[DEBUG] MicrosoftAppId: {AppId}", string.IsNullOrEmpty(appId) ? "<empty>" : appId);
    //logger.LogInformation("[DEBUG] MicrosoftAppPassword: {AppPassword}", string.IsNullOrEmpty(appPassword) ? "<empty>" : new string('*', Math.Min(8, appPassword!.Length)));
    //logger.LogInformation("[DEBUG] MicrosoftAppTenantId: {AppTenantId}", string.IsNullOrEmpty(tenantId) ? "<empty>" : tenantId);
    //logger.LogInformation("[DEBUG] MicrosoftAppType: {AppType}", string.IsNullOrEmpty(appType) ? "<empty>" : appType);

    if (!credsProvided)
    {
        // ローカル開発（Emulator）向け: 認証なしで起動
        logger.LogWarning("[DEBUG] AppId/Password が未設定のため、ローカル用に Authentication を無効化して起動します（Emulator からの匿名アクセスを許可）。");
    }

    // Log detailed authentication request information（実際に資格情報がある場合のみ詳細を出す）
    if (credsProvided)
    {
        var effectiveAppType = string.IsNullOrWhiteSpace(appType) ? "SingleTenant" : appType;
        logger.LogInformation("[DEBUG] Preparing authentication request with the following details:");
        logger.LogInformation("[DEBUG] Authentication Endpoint: https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token", tenantId);
        logger.LogInformation("[DEBUG] MicrosoftAppId: {AppId}", appId);
        logger.LogInformation("[DEBUG] MicrosoftAppTenantId: {TenantId}", tenantId);
        logger.LogInformation("[DEBUG] MicrosoftAppType: {AppType}", effectiveAppType);
    }

    // 構成を条件付きで組み立て
    var settings = new Dictionary<string, string?>();
    if (credsProvided)
    {
        var effectiveAppType = string.IsNullOrWhiteSpace(appType) ? "SingleTenant" : appType;
        settings["MicrosoftAppId"] = appId;
        settings["MicrosoftAppPassword"] = appPassword;
        settings["MicrosoftAppType"] = effectiveAppType;
        if (string.Equals(effectiveAppType, "SingleTenant", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(tenantId))
            {
                logger.LogWarning("[DEBUG] MicrosoftAppTenantId が未設定ですが、AppType=SingleTenant です。認証に失敗します。Azure での動作用にテナントIDを設定してください。");
            }
            else
            {
                settings["MicrosoftAppTenantId"] = tenantId;
            }
        }
    }
    // credsProvided=false の場合は settings を空のまま渡す → SDK が Authentication Disabled として動作

    var botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(new ConfigurationBuilder()
        .AddInMemoryCollection(settings)
        .Build());

    return new CloudAdapter(botFrameworkAuthentication, logger);
});
// SimpleBot を DI へ登録（OneDriveExcelService, Kernel, ConversationState をオプション注入）
builder.Services.AddTransient<IBot>(sp =>
{
    var kernel = sp.GetService<Kernel>();
    var convo = sp.GetService<ConversationState>();
    var oneDrive = sp.GetService<PhraseXross.Services.OneDriveExcelService>();
    return new SimpleBot(kernel, convo, oneDrive);
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

    Console.WriteLine($"[DEBUG] AOAI_ENDPOINT: {endpoint}");
    Console.WriteLine($"[DEBUG] AOAI_API_KEY: {apiKey}");
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
