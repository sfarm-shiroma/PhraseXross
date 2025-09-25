using Azure.Core;
using Azure.Identity;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Linq;

namespace PhraseXross.Services;

public class OneDriveExcelService
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<OneDriveExcelService> _logger;
    private readonly IConfiguration _config;

    public OneDriveExcelService(IHttpClientFactory httpClientFactory, ILogger<OneDriveExcelService> logger, IConfiguration config)
    {
        _httpClientFactory = httpClientFactory;
        _logger = logger;
        _config = config;
    }

    // uploadedCallback: アップロード完了直後（セル書き込み前）に WebUrl を通知
    public async Task<OneDriveUploadResult> CreateAndFillExcelAsync(IProgress<string>? uploadedCallback = null, CancellationToken ct = default)
    {
    // 優先順: Bot 用既存キー (MicrosoftAppId / MicrosoftAppTenantId) → OneDrive 専用キー (新 / 旧) → 構成キー
    var clientId = Environment.GetEnvironmentVariable("MicrosoftAppId")
               ?? _config["MicrosoftAppId"]
               ?? Environment.GetEnvironmentVariable("OneDriveClientId")
               ?? Environment.GetEnvironmentVariable("ONEDRIVE_CLIENT_ID")
               ?? _config["OneDriveClientId"]
               ?? _config["OneDrive:ClientId"]
               ?? string.Empty;
    var tenantId = Environment.GetEnvironmentVariable("MicrosoftAppTenantId")
               ?? _config["MicrosoftAppTenantId"]
               ?? Environment.GetEnvironmentVariable("OneDriveTenantId")
               ?? Environment.GetEnvironmentVariable("ONEDRIVE_TENANT_ID")
               ?? _config["OneDriveTenantId"]
               ?? _config["OneDrive:TenantId"]
               ?? string.Empty;
        if (string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(tenantId))
        {
            // 例外ではなく失敗結果を返却して上位でユーザーフレンドリーな文言を出す
            return OneDriveUploadResult.Fail("OneDriveの資格情報が未設定です。MicrosoftAppId / MicrosoftAppTenantId もしくは OneDriveClientId / OneDriveTenantId（旧: ONEDRIVE_CLIENT_ID / ONEDRIVE_TENANT_ID）のいずれかを設定してください。");
        }
        var scopes = new[] { "Files.ReadWrite" };

        var credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions
        {
            ClientId = clientId,
            TenantId = tenantId,
            DeviceCodeCallback = (code, cancellationToken) =>
            {
                _logger.LogInformation("Device code: {Message}", code.Message);
                return Task.CompletedTask;
            }
        });

        var token = await credential.GetTokenAsync(new TokenRequestContext(scopes), ct);
        _logger.LogInformation("Graph token acquired (expires {ExpiresOn})", token.ExpiresOn);

        var http = _httpClientFactory.CreateClient("graph");
        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        // copy template
        var template = Path.Combine(AppContext.BaseDirectory, "Templates", "EmptyExcel.xlsx");
        if (!File.Exists(template))
        {
            throw new FileNotFoundException("Template not found", template);
        }
        var stamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
        var fileName = stamp + ".xlsx";
        var tempPath = Path.Combine(Path.GetTempPath(), fileName);
        File.Copy(template, tempPath, true);

        try
        {
            // drive id
            var driveRes = await SendWithRetry(() => http.GetAsync("https://graph.microsoft.com/v1.0/me/drive", ct), ct);
            if (!driveRes.IsSuccessStatusCode)
            {
                return OneDriveUploadResult.Fail($"Drive lookup failed {driveRes.StatusCode}: {await driveRes.Content.ReadAsStringAsync(ct)}");
            }
            var driveJson = await driveRes.Content.ReadAsStringAsync(ct);
            using var driveDoc = System.Text.Json.JsonDocument.Parse(driveJson);
            var driveId = driveDoc.RootElement.GetProperty("id").GetString();
            if (string.IsNullOrEmpty(driveId)) return OneDriveUploadResult.Fail("DriveId missing");

            await using (var fs = File.OpenRead(tempPath))
            {
                var uploadUrl = $"https://graph.microsoft.com/v1.0/drives/{driveId}/root:/{fileName}:/content";
                var putRes = await SendWithRetry(() => http.PutAsync(uploadUrl, new StreamContent(fs), ct), ct);
                if (!putRes.IsSuccessStatusCode)
                {
                    return OneDriveUploadResult.Fail($"Upload failed {putRes.StatusCode}: {await putRes.Content.ReadAsStringAsync(ct)}");
                }
                var upJson = await putRes.Content.ReadAsStringAsync(ct);
                using var upDoc = System.Text.Json.JsonDocument.Parse(upJson);
                var itemId = upDoc.RootElement.GetProperty("id").GetString();
                var webUrl = upDoc.RootElement.GetProperty("webUrl").GetString();
                if (itemId == null || webUrl == null) return OneDriveUploadResult.Fail("Upload response incomplete");

                // アップロード直後に URL をコールバック（セル書き込み前）
                if (webUrl == null)
                {
                    return OneDriveUploadResult.Fail("webUrl not returned from Graph upload response");
                }
                try { uploadedCallback?.Report(webUrl); } catch { /* ignore */ }

                // 指定のシート（状況 / 課題・欲求 / 感情 / 温度感）を追加作成
                var desiredSheets = new[] { "状況", "課題・欲求", "感情", "温度感" };
                foreach (var sheetName in desiredSheets)
                {
                    try
                    {
                        var addUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets";
                        var body = $"{{\"name\":\"{sheetName}\"}}";
                        using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                        var addRes = await SendWithRetry(() => http.PostAsync(addUrl, content, ct), ct);
                        if (!addRes.IsSuccessStatusCode)
                        {
                            var respText = await addRes.Content.ReadAsStringAsync(ct);
                            _logger.LogWarning("Worksheet create failed ({Sheet}) {Status}: {Resp}", sheetName, addRes.StatusCode, respText);
                        }
                        else
                        {
                            _logger.LogInformation("Worksheet '{Sheet}' created", sheetName);
                        }
                    }
                    catch (Exception exSheet)
                    {
                        _logger.LogWarning(exSheet, "Worksheet create exception for {Sheet}", sheetName);
                    }
                }

                return OneDriveUploadResult.CreateSuccess(webUrl, fileName);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Upload process failed");
            return OneDriveUploadResult.Fail(ex.Message);
        }
        finally
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
        }
    }

    private static async Task<HttpResponseMessage> SendWithRetry(Func<Task<HttpResponseMessage>> action, CancellationToken ct)
    {
        const int max = 5;
        int count = 0;
        while (true)
        {
            var res = await action();
            if (res.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
            {
                count++;
                if (count > max) throw new Exception("Rate limit exceeded retry attempts");
                int delayMs = 1000;
                if (res.Headers.TryGetValues("Retry-After", out var vals) && int.TryParse(vals.FirstOrDefault(), out int sec))
                {
                    delayMs = sec * 1000;
                }
                await Task.Delay(delayMs, ct);
                continue;
            }
            return res;
        }
    }
}

public record OneDriveUploadResult(bool IsSuccess, string? WebUrl, string? FileName, string? Error)
{
    public static OneDriveUploadResult CreateSuccess(string webUrl, string fileName) => new(true, webUrl, fileName, null);
    public static OneDriveUploadResult Fail(string error) => new(false, null, null, error);
}
