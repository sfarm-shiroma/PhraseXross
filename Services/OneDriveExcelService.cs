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

                // ========= 動的シート生成: Step6 JSON からカテゴリ抽出 → 単体シート + 全てのペアシート =========
                // 1) 最新 creative_elements_*.json を探す
                var exportDir = Path.Combine(AppContext.BaseDirectory, "exports");
                string? latestJson = null;
                try
                {
                    if (Directory.Exists(exportDir))
                    {
                        latestJson = Directory.GetFiles(exportDir, "creative_elements_*.json")
                            .OrderByDescending(f => f) // ファイル名にUTCタイムスタンプ yyyyMMdd_HHmmss を含むため文字列降順で最新
                            .FirstOrDefault();
                    }
                }
                catch (Exception exList)
                {
                    _logger.LogWarning(exList, "Failed to enumerate creative elements JSON files");
                }

                List<string> categoryKeys = new();
                if (!string.IsNullOrWhiteSpace(latestJson) && File.Exists(latestJson))
                {
                    try
                    {
                        using var fsJson = File.OpenRead(latestJson);
                        using var doc = System.Text.Json.JsonDocument.Parse(fsJson);
                        // JSON オブジェクト直下のプロパティ名を列挙順で保持
                        foreach (var prop in doc.RootElement.EnumerateObject())
                        {
                            if (prop.Value.ValueKind == System.Text.Json.JsonValueKind.Array)
                            {
                                categoryKeys.Add(prop.Name.Trim());
                            }
                        }
                        _logger.LogInformation("Creative elements categories detected: {Cats}", string.Join(",", categoryKeys));
                    }
                    catch (Exception exParse)
                    {
                        _logger.LogWarning(exParse, "Failed to parse creative elements JSON. Fallback to default categories.");
                    }
                }

                // 以前は JSON 無しの場合に従来4カテゴリ（状況/課題・欲求/感情/温度感）を追加していたが、要求により削除。
                if (categoryKeys.Count == 0)
                {
                    _logger.LogInformation("No creative element categories found in latest JSON. Skipping worksheet creation.");
                }

                // Excel シート名制約: 長さ <=31, 特定記号 : \\ / ? * [ ] 不可
                string SanitizeSheetName(string raw)
                {
                    var invalid = new char[] { ':', '\\', '/', '?', '*', '[', ']' };
                    var cleaned = new string(raw.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
                    if (cleaned.Length > 31) cleaned = cleaned.Substring(0, 31);
                    return string.IsNullOrWhiteSpace(cleaned) ? "Sheet" : cleaned;
                }

                async Task CreateSheetIfNotExistsAsync(string sheetName)
                {
                    var safe = SanitizeSheetName(sheetName);
                    try
                    {
                        var addUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets";
                        var body = $"{{\"name\":\"{safe}\"}}";
                        using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                        var addRes = await SendWithRetry(() => http.PostAsync(addUrl, content, ct), ct);
                        if (!addRes.IsSuccessStatusCode)
                        {
                            var respText = await addRes.Content.ReadAsStringAsync(ct);
                            // すでに存在する場合 400 / 409 の可能性 → 情報レベルで残す
                            _logger.LogWarning("Worksheet create failed ({Sheet}) {Status}: {Resp}", safe, addRes.StatusCode, respText);
                        }
                        else
                        {
                            _logger.LogInformation("Worksheet '{Sheet}' created", safe);
                        }
                    }
                    catch (Exception exSheet)
                    {
                        _logger.LogWarning(exSheet, "Worksheet create exception for {Sheet}", safe);
                    }
                }

                var created = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (categoryKeys.Count > 0)
                {
                    // 要望: 単体カテゴリシートは作らず、既定の "Sheet1" を最初のペア名にリネームし残りのみ追加
                    string? firstPairName = null;
                    if (categoryKeys.Count >= 2)
                    {
                        firstPairName = categoryKeys[0] + "×" + categoryKeys[1];
                        // 既存シート一覧取得
                        try
                        {
                            var listRes = await SendWithRetry(() => http.GetAsync($"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets", ct), ct);
                            if (listRes.IsSuccessStatusCode)
                            {
                                var listJson = await listRes.Content.ReadAsStringAsync(ct);
                                using var listDoc = System.Text.Json.JsonDocument.Parse(listJson);
                                if (listDoc.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == System.Text.Json.JsonValueKind.Array)
                                {
                                    var sheetNames = new List<string>();
                                    foreach (var e in arr.EnumerateArray())
                                    {
                                        if (e.TryGetProperty("name", out var nmEl))
                                        {
                                            sheetNames.Add(nmEl.GetString() ?? "(null)");
                                        }
                                    }
                                    _logger.LogInformation("Existing workbook sheets before rename attempt: {Sheets}", string.Join(",", sheetNames));
                                    var sheet1 = arr.EnumerateArray()
                                        .FirstOrDefault(e => e.TryGetProperty("name", out var nm) && nm.GetString() != null && string.Equals(nm.GetString(), "Sheet1", StringComparison.OrdinalIgnoreCase));
                                    if (sheet1.ValueKind == System.Text.Json.JsonValueKind.Object && sheet1.TryGetProperty("id", out var idEl))
                                    {
                                        var wsId = idEl.GetString();
                                        if (!string.IsNullOrWhiteSpace(wsId))
                                        {
                                            var safeName = SanitizeSheetName(firstPairName);
                                            var patchUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{wsId}";
                                            var bodyJson = $"{{\"name\":\"{safeName}\"}}"; // 正しいJSON文字列
                                            var patchRes = await SendWithRetry(() => {
                                                var req = new HttpRequestMessage(new HttpMethod("PATCH"), patchUrl);
                                                req.Content = new StringContent(bodyJson, System.Text.Encoding.UTF8, "application/json");
                                                return http.SendAsync(req, ct);
                                            }, ct);
                                            if (patchRes.IsSuccessStatusCode)
                                            {
                                                _logger.LogInformation("Renamed Sheet1 to {NewName}", safeName);
                                                created.Add(firstPairName); // マークして後続生成をスキップ
                                            }
                                            else
                                            {
                                                var respT = await patchRes.Content.ReadAsStringAsync(ct);
                                                _logger.LogWarning("Failed to rename Sheet1 to {NewName} {Status}: {Resp}", safeName, patchRes.StatusCode, respT);
                                                // 失敗時は後で通常作成を試みるため created に追加しない
                                            }
                                        }
                                    }
                                    else
                                    {
                                        _logger.LogInformation("Sheet1 not found by name in workbook (names: {Sheets}). Trying direct name-based PATCH fallback.", string.Join(",", sheetNames));
                                        try
                                        {
                                            var safeName = SanitizeSheetName(firstPairName);
                                            var directPatchUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/Sheet1"; // 名前指定ルート
                                            var bodyJson = $"{{\"name\":\"{safeName}\"}}";
                                            var patchRes = await SendWithRetry(() => {
                                                var req = new HttpRequestMessage(new HttpMethod("PATCH"), directPatchUrl);
                                                req.Content = new StringContent(bodyJson, System.Text.Encoding.UTF8, "application/json");
                                                return http.SendAsync(req, ct);
                                            }, ct);
                                            if (patchRes.IsSuccessStatusCode)
                                            {
                                                _logger.LogInformation("(Fallback) Renamed Sheet1 to {NewName}", safeName);
                                                created.Add(firstPairName);
                                            }
                                            else
                                            {
                                                var respT = await patchRes.Content.ReadAsStringAsync(ct);
                                                _logger.LogWarning("(Fallback) Failed to rename Sheet1 to {NewName} {Status}: {Resp}", safeName, patchRes.StatusCode, respT);
                                            }
                                        }
                                        catch (Exception exFallback)
                                        {
                                            _logger.LogWarning(exFallback, "Fallback rename attempt failed");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                _logger.LogWarning("Worksheet list fetch failed {Status}", listRes.StatusCode);
                            }
                        }
                        catch (Exception exListWs)
                        {
                            _logger.LogWarning(exListWs, "Worksheet list / rename process failed");
                        }
                    }

                    // 残りペア（最初のペアはリネーム済みならスキップ）
                    for (int i = 0; i < categoryKeys.Count; i++)
                    {
                        for (int j = i + 1; j < categoryKeys.Count; j++)
                        {
                            var pairName = categoryKeys[i] + "×" + categoryKeys[j];
                            if (created.Contains(pairName)) continue; // 既にSheet1をリネーム済み
                            if (created.Add(pairName))
                                await CreateSheetIfNotExistsAsync(pairName);
                        }
                    }
                }
                _logger.LogInformation("Worksheet generation complete. Total attempted: {Count}", created.Count);

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
