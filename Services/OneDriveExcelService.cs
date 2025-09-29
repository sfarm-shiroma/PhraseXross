using Azure.Core;
using Azure.Identity;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Linq;
// Semantic Kernel は本サービスでは使用せず、Azure OpenAI 環境変数を直接利用して呼び出す方針

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
    // delegatedAccessToken: Teams SSO / OAuth 経由で既に取得済みのユーザー委譲トークン（preferred）。null の場合は開発用フォールバックとして DeviceCode Flow を試行。
    public async Task<OneDriveUploadResult> CreateAndFillExcelAsync(
        IProgress<string>? uploadedCallback = null,
        string? taglineSummaryJson = null,
        CancellationToken ct = default,
        string? delegatedAccessToken = null)
    {
        if (string.IsNullOrWhiteSpace(delegatedAccessToken))
        {
            // フォールバック（DeviceCode）は本番要件と乖離するため廃止。呼び出し側で SignInCard を提示する運用。
            return OneDriveUploadResult.Fail("ユーザーの委譲トークンが未取得のため Excel 出力を実行できません。まずサインインしてください。");
        }

        var accessToken = delegatedAccessToken!;
        _logger.LogDebug("Using delegated Graph access token (length={Len}).", accessToken.Length);

        var http = _httpClientFactory.CreateClient("graph");
        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

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

                // ========= 動的シート生成: Step7 クリエイティブ要素 JSON からカテゴリ抽出 → 全ペアシート =========
                // クリエイティブ要素ファイル探索
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

                // 生成: カテゴリ名と要素配列を取得
                List<string> categoryKeys = new();
                Dictionary<string, List<string>> categoryValues = new();
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
                                var name = prop.Name.Trim();
                                categoryKeys.Add(name);
                                var list = new List<string>();
                                foreach (var el in prop.Value.EnumerateArray())
                                {
                                    if (el.ValueKind == System.Text.Json.JsonValueKind.String)
                                    {
                                        var v = el.GetString();
                                        if (!string.IsNullOrWhiteSpace(v)) list.Add(v!.Trim());
                                    }
                                }
                                categoryValues[name] = list;
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

                // 要素書き込み: A2 縦に firstCat の値, B2 から横方向に secondCat の値
                async Task FillPairSheetAsync(string sheetName, string firstCat, string secondCat)
                {
                    var safe = SanitizeSheetName(sheetName);
                    if (!categoryValues.TryGetValue(firstCat, out var firstList) || firstList.Count == 0) { _logger.LogWarning("No values for category {Cat}", firstCat); return; }
                    if (!categoryValues.TryGetValue(secondCat, out var secondList) || secondList.Count == 0) { _logger.LogWarning("No values for category {Cat}", secondCat); return; }

                    string ColLetter(int index1Based)
                    {
                        var dividend = index1Based;
                        var col = string.Empty;
                        while (dividend > 0)
                        {
                            var modulo = (dividend - 1) % 26;
                            col = Convert.ToChar('A' + modulo) + col;
                            dividend = (dividend - modulo) / 26;
                        }
                        return col;
                    }

                    // ヘッダー/ボディ配置仕様:
                    // A1: 空白
                    // B1..: secondCat の要素（横）
                    // A2..: firstCat の要素（縦）
                    // B2..: 交差セル  firstItem × secondItem （要望 #2）

                    // 縦: A2..A(1+count)
                    int vCount = firstList.Count;
                    int lastRow = 1 + vCount; // 2開始 → 2..(1+count)
                    var vertAddress = $"A2:A{lastRow}";
                    var vertValues = firstList.Select(v => new object[] { v }).ToList();
                    var vertBody = new { values = vertValues };
                    var vertJson = System.Text.Json.JsonSerializer.Serialize(vertBody);
                    var vertUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{vertAddress}')";
                    try
                    {
                        using var content = new StringContent(vertJson, System.Text.Encoding.UTF8, "application/json");
                        var res = await SendWithRetry(() => http.PatchAsync(vertUrl, content, ct), ct);
                        if (!res.IsSuccessStatusCode)
                        {
                            _logger.LogWarning("Vertical write failed {Sheet} {Status}", safe, res.StatusCode);
                        }
                    }
                    catch (Exception exV)
                    {
                        _logger.LogWarning(exV, "Vertical write exception {Sheet}", safe);
                    }

                    // 横ヘッダー: B1..(column)1  (要素数 n → B..(B+n-1))
                    int hCount = secondList.Count;
                    int lastColIndex = 2 + hCount - 1; // B=2
                    var lastColLetter = ColLetter(lastColIndex);
                    var horizAddress = $"B1:{lastColLetter}1";
                    var horizValues = new List<object[]> { secondList.Cast<object>().ToArray() };
                    var horizBody = new { values = horizValues };
                    var horizJson = System.Text.Json.JsonSerializer.Serialize(horizBody);
                    var horizUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{horizAddress}')";
                    try
                    {
                        using var content = new StringContent(horizJson, System.Text.Encoding.UTF8, "application/json");
                        var res = await SendWithRetry(() => http.PatchAsync(horizUrl, content, ct), ct);
                        if (!res.IsSuccessStatusCode)
                        {
                            _logger.LogWarning("Horizontal write failed {Sheet} {Status}", safe, res.StatusCode);
                        }
                    }
                    catch (Exception exH)
                    {
                        _logger.LogWarning(exH, "Horizontal write exception {Sheet}", safe);
                    }

                    // 軸セルの塗りつぶし + 太字化: 縦軸(#FBE2D5), 横軸(#DAE9F8)
                    async Task ApplyFillAsync(string address, string color)
                    {
                        var fillUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{address}')/format/fill";
                        var body = $"{{\"color\":\"{color}\"}}"; // { "color": "#XXXXXX" }
                        try
                        {
                            using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                            var res = await SendWithRetry(() => http.PatchAsync(fillUrl, content, ct), ct);
                            if (!res.IsSuccessStatusCode)
                            {
                                _logger.LogWarning("Fill color failed {Sheet} {Addr} {Status}", safe, address, res.StatusCode);
                            }
                        }
                        catch (Exception exF)
                        {
                            _logger.LogWarning(exF, "Fill color exception {Sheet} {Addr}", safe, address);
                        }
                    }

                    // 軸セルフォント太字化
                    async Task ApplyBoldAsync(string address)
                    {
                        var fontUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{address}')/format/font";
                        var body = "{\"bold\":true}"; // { "bold": true }
                        try
                        {
                            using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                            var res = await SendWithRetry(() => http.PatchAsync(fontUrl, content, ct), ct);
                            if (!res.IsSuccessStatusCode)
                            {
                                _logger.LogWarning("Set bold failed {Sheet} {Addr} {Status}", safe, address, res.StatusCode);
                            }
                        }
                        catch (Exception exB)
                        {
                            _logger.LogWarning(exB, "Set bold exception {Sheet} {Addr}", safe, address);
                        }
                    }

                    // A列 (A2..A{lastRow}) と 横ヘッダー (B1..{lastColLetter}1)
                    await ApplyFillAsync($"A2:A{lastRow}", "#FBE2D5");
                    await ApplyFillAsync(horizAddress, "#DAE9F8");
                    // 太字化（縦軸+横軸）
                    await ApplyBoldAsync($"A2:A{lastRow}");
                    await ApplyBoldAsync(horizAddress);

                    // 列幅調整: A ~ lastColLetter を幅 300pt に統一（可視性優先。必要なら環境変数化予定）
                    async Task SetColumnWidthAsync(string fromCol, string toCol, double width)
                    {
                        var absRange = fromCol == toCol ? $"${fromCol}:${toCol}" : $"${fromCol}:${toCol}"; // 列全体 ($A:$D)
                        var body = $"{{\"columnWidth\":{width.ToString(System.Globalization.CultureInfo.InvariantCulture)} }}";
                        async Task<bool> TryPatchAsync(string rangeAddress)
                        {
                            var url = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{rangeAddress}')/format";
                            using var content = new StringContent(body, System.Text.Encoding.UTF8, "application/json");
                            var res = await SendWithRetry(() => http.PatchAsync(url, content, ct), ct);
                            if (!res.IsSuccessStatusCode)
                            {
                                var resp = await res.Content.ReadAsStringAsync(ct);
                                _logger.LogWarning("Set column width failed {Sheet} {Range} {Status} {Resp}", safe, rangeAddress, res.StatusCode, resp);
                                return false;
                            }
                            return true;
                        }
                        try
                        {
                            // 1) 列全体指定 ($A:$D) を試す
                            if (!await TryPatchAsync(absRange))
                            {
                                // 2) フォールバック: 1行目セル範囲 ($A$1:$D$1)
                                var row1Range = fromCol == toCol ? $"${fromCol}$1:${toCol}$1" : $"${fromCol}$1:${toCol}$1";
                                await TryPatchAsync(row1Range);
                            }
                        }
                        catch (Exception exW)
                        {
                            _logger.LogWarning(exW, "Set column width exception {Sheet} {From}-{To}", safe, fromCol, toCol);
                        }
                    }
                    await SetColumnWidthAsync("A", lastColLetter, 300); // 300pt に拡大

                    // 設定結果の一部を取得しログ（A列）
                    try
                    {
                        var getFormatUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='$A:$A')/format";
                        var getRes = await SendWithRetry(() => http.GetAsync(getFormatUrl, ct), ct);
                        if (getRes.IsSuccessStatusCode)
                        {
                            var txt = await getRes.Content.ReadAsStringAsync(ct);
                            using var fmtDoc = System.Text.Json.JsonDocument.Parse(txt);
                            if (fmtDoc.RootElement.TryGetProperty("columnWidth", out var wEl) && wEl.ValueKind == System.Text.Json.JsonValueKind.Number)
                            {
                                _logger.LogInformation("Column A width reported by Graph: {WidthPt}pt", wEl.GetDouble());
                            }
                            else
                            {
                                _logger.LogInformation("Column A width GET succeeded but columnWidth not present: {Json}", txt);
                            }
                        }
                        else
                        {
                            _logger.LogWarning("Column A width GET failed {Status}", getRes.StatusCode);
                        }
                    }
                    catch (Exception exGW)
                    {
                        _logger.LogWarning(exGW, "Failed to get column width after setting");
                    }

                    // 交差セル: 生成AIに単発プロンプト送信し結果(5行)を書き込む（履歴非蓄積）
                    // 既存課題: 完了後に一括 PATCH していたため、途中でプロセス終了すると Excel へ反映されない → 行単位で逐次反映
                    try
                    {
                        var flushMode = Environment.GetEnvironmentVariable("TAGLINE_FLUSH_MODE")?.Trim().ToUpperInvariant();
                        if (string.IsNullOrWhiteSpace(flushMode)) flushMode = "ROW"; // ROW | ALL （将来: CELL）
                        var matrixAddressAll = $"B2:{lastColLetter}{lastRow}";
                        var matrixValuesAll = new List<object[]>();
                        // AI エンドポイント設定（任意）。未設定ならプロンプトそのものをセルへ。
                        // 優先順: TAGLINE_AI_* → AOAI_*
                        var taglineEndpoint = Environment.GetEnvironmentVariable("TAGLINE_AI_ENDPOINT") ?? _config["TaglineAI:Endpoint"];
                        var taglineKey = Environment.GetEnvironmentVariable("TAGLINE_AI_KEY") ?? _config["TaglineAI:Key"];
                        var aoaiEndpoint = Environment.GetEnvironmentVariable("AOAI_ENDPOINT");
                        var aoaiKey = Environment.GetEnvironmentVariable("AOAI_API_KEY");
                        var aoaiDeployment = Environment.GetEnvironmentVariable("AOAI_DEPLOYMENT");
                        var aoaiApiVersion = Environment.GetEnvironmentVariable("AOAI_API_VERSION") ?? "2024-05-01-preview";

                        bool azureOpenAiAvailable = string.IsNullOrWhiteSpace(taglineEndpoint) && !string.IsNullOrWhiteSpace(aoaiEndpoint) && !string.IsNullOrWhiteSpace(aoaiKey) && !string.IsNullOrWhiteSpace(aoaiDeployment);
                        bool simpleHttpAvailable = !string.IsNullOrWhiteSpace(taglineEndpoint) && !string.IsNullOrWhiteSpace(taglineKey);
                        for (int r = 0; r < vCount; r++)
                        {
                            var row = new object[hCount];
                            for (int c = 0; c < hCount; c++)
                            {
                                var parentVertical = firstCat;
                                var parentHorizontal = secondCat;
                                var childVertical = firstList[r];
                                var childHorizontal = secondList[c];
                                string prompt = BuildTaglinePrompt(taglineSummaryJson, parentVertical, childVertical, parentHorizontal, childHorizontal);
                                string cellValue;
                                if (azureOpenAiAvailable || simpleHttpAvailable)
                                {
                                    try
                                    {
                                        // セルアドレス計算 (B2 起点)
                                        string CellAddress()
                                        {
                                            var colLetter = ColLetter(2 + c); // B=2
                                            int rowNum = 2 + r; // 行基準 2
                                            return colLetter + rowNum.ToString();
                                        }
                                        var addr = CellAddress();
                                        var promptPreview = prompt.Length > 240 ? prompt.Substring(0, 240) + "..." : prompt;
                                        _logger.LogInformation("[TaglineGen][START] sheet={Sheet} cell={Cell} promptChars={Len} via={Mode} preview=\n{Preview}", safe, addr, prompt.Length, azureOpenAiAvailable?"AzureOpenAI":"SimpleHTTP", promptPreview);
                                        var started = DateTime.UtcNow;
                                        if (azureOpenAiAvailable)
                                        {
                                            cellValue = await GenerateTaglinesAzureAsync(aoaiEndpoint!, aoaiDeployment!, aoaiKey!, aoaiApiVersion, prompt, ct);
                                        }
                                        else // simpleHttpAvailable
                                        {
                                            cellValue = await GenerateTaglinesSimpleAsync(taglineEndpoint!, taglineKey!, prompt, ct);
                                        }
                                        var elapsed = (DateTime.UtcNow - started).TotalMilliseconds;
                                        // 行数と先頭行（秘匿情報なし想定）をログ
                                        var lines = cellValue.Replace("\r", "").Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
                                        var firstLine = lines.Count > 0 ? lines[0] : "(empty)";
                                        _logger.LogInformation("[TaglineGen][DONE] sheet={Sheet} cell={Cell} ms={Ms:F0} lines={Count} first='{First}'", safe, addr, elapsed, lines.Count, firstLine);
                                    }
                                    catch (Exception exGen)
                                    {
                                        _logger.LogWarning(exGen, "Tagline generation failed; fallback to prompt text");
                                        cellValue = prompt; // フォールバック: プロンプトそのもの
                                    }
                                    // レート制御 (簡易) 1秒待機 (SKでも過負荷防止のため統一)
                                    try { await Task.Delay(1000, ct); } catch { }
                                }
                                else
                                {
                                    // AI未設定 → プロンプト表示のみ
                                    cellValue = prompt;
                                }
                                row[c] = cellValue;
                            }
                            if (flushMode == "ROW")
                            {
                                // 行単位で即時反映
                                var rowIndex1Based = 2 + r; // B2 開始
                                var rowAddress = $"B{rowIndex1Based}:{lastColLetter}{rowIndex1Based}";
                                var body = new { values = new List<object[]> { row } };
                                var bodyJson = System.Text.Json.JsonSerializer.Serialize(body);
                                var rowUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{rowAddress}')";
                                try
                                {
                                    using var content = new StringContent(bodyJson, System.Text.Encoding.UTF8, "application/json");
                                    var res = await SendWithRetry(() => http.PatchAsync(rowUrl, content, ct), ct);
                                    if (!res.IsSuccessStatusCode)
                                    {
                                        _logger.LogWarning("Matrix row write failed {Sheet} row={Row} {Status}", safe, rowIndex1Based, res.StatusCode);
                                    }
                                }
                                catch (Exception exRow)
                                {
                                    _logger.LogWarning(exRow, "Matrix row write exception {Sheet} row={Row}", safe, rowIndex1Based);
                                }
                            }
                            else // ALL モード
                            {
                                matrixValuesAll.Add(row);
                            }
                        }
                        if (flushMode == "ALL")
                        {
                            var matrixBody = new { values = matrixValuesAll };
                            var matrixJson = System.Text.Json.JsonSerializer.Serialize(matrixBody);
                            var matrixUrl = $"https://graph.microsoft.com/v1.0/me/drive/items/{itemId}/workbook/worksheets/{Uri.EscapeDataString(safe)}/range(address='{matrixAddressAll}')";
                            using var content = new StringContent(matrixJson, System.Text.Encoding.UTF8, "application/json");
                            var res = await SendWithRetry(() => http.PatchAsync(matrixUrl, content, ct), ct);
                            if (!res.IsSuccessStatusCode)
                            {
                                _logger.LogWarning("Matrix prompt write failed {Sheet} {Status}", safe, res.StatusCode);
                            }
                        }
                    }
                    catch (Exception exM)
                    {
                        _logger.LogWarning(exM, "Matrix prompt write exception {Sheet}", safe);
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
                                                // リネーム成功後に要素書き込み
                                                await FillPairSheetAsync(firstPairName, categoryKeys[0], categoryKeys[1]);
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
                                                await FillPairSheetAsync(firstPairName, categoryKeys[0], categoryKeys[1]);
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
                            {
                                await CreateSheetIfNotExistsAsync(pairName);
                                await FillPairSheetAsync(pairName, categoryKeys[i], categoryKeys[j]);
                            }
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

    // キャッチフレーズ生成用プロンプト組み立て
    // summaryJson: Step6 要約 JSON（補助情報）
    // parentVertical / parentHorizontal: カテゴリ名
    // childVertical / childHorizontal: それぞれの具体的要素
    // 要望: "[親カテゴリと子要素]" の 2 行（状況 / 課題・欲求 など）をメイン焦点として連想し、
    // summaryJson は補助的な背景として扱うようプロンプトを再構成。
    private static string BuildTaglinePrompt(string? summaryJson, string parentVertical, string childVertical, string parentHorizontal, string childHorizontal)
    {
        // 余計な改行を圧縮 + Unicodeエスケープ \uXXXX をデコード（AI応答視認性向上目的）
        string DecodeUnicode(string s)
        {
            return System.Text.RegularExpressions.Regex.Replace(s, @"\\u([0-9A-Fa-f]{4})", m =>
            {
                var code = Convert.ToInt32(m.Groups[1].Value, 16);
                return char.ConvertFromUtf32(code);
            });
        }
        string Clean(string? t)
        {
            if (string.IsNullOrWhiteSpace(t)) return "";
            var r = t.Replace("\r", "").Trim();
            try { r = DecodeUnicode(r); } catch { /* ignore decode errors */ }
            return r;
        }
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("あなたは日本語のプロフェッショナルコピーライターです。");
        sb.AppendLine("最優先で連想・統合すべきメイン情報は直後の '[メインペア]' です。以降の '[参考情報]' は補助的に使い、メインの二つの行から核心を抽出・掛け合わせたインパクトと一貫性を重視してください。");
        sb.AppendLine();
        sb.AppendLine("[メインペア]");
        sb.AppendLine($"{parentVertical}:{childVertical}");
        sb.AppendLine($"{parentHorizontal}:{childHorizontal}");
        sb.AppendLine();
        if (!string.IsNullOrWhiteSpace(summaryJson))
        {
            sb.AppendLine("[参考情報 JSON 要約] ※必要な要素のみ暗黙的に活かし、丸写しや列挙をしない");
            sb.AppendLine(Clean(summaryJson));
            sb.AppendLine();
        }
        sb.AppendLine("[生成指針]");
        sb.AppendLine("- メインペアのシナジー・緊張関係・ベネフィットを核にする");
        sb.AppendLine("- 参考情報は語調/差別化/独自性強化のヒントとしてのみ利用");
        sb.AppendLine("- JSONのフィールド名や記号、括弧、コロンなどは出力に含めない");
        sb.AppendLine("- 誇大/医療的/裏付けの無い断定を避ける");
        sb.AppendLine();
        sb.AppendLine("[出力フォーマット]");
        sb.AppendLine("キャッチフレーズのみを改行区切りで5行。番号・記号・引用符・前後余白・解説禁止。");
        sb.AppendLine();
        sb.AppendLine("[禁止事項]");
        sb.AppendLine("- 同じ語尾や同じ語の繰り返しで単調になること");
        sb.AppendLine("- メインペア文の丸写し");
        sb.AppendLine("- JSONそのもの/キー名の再掲");
        sb.AppendLine();
        sb.AppendLine("出力開始→");
        return sb.ToString();
    }

    // SK 利用パス: Kernel が存在する場合はこちらで生成
    // シンプルHTTP (ユーザー独自エンドポイント) 呼び出し
    private static async Task<string> GenerateTaglinesSimpleAsync(string endpoint, string apiKey, string prompt, CancellationToken ct)
    {
        using var http = new HttpClient();
        http.DefaultRequestHeaders.Add("Authorization", "Bearer " + apiKey);
        var payload = new { prompt = prompt }; // エンドポイント側で 'prompt' を受け取る想定
        var json = System.Text.Json.JsonSerializer.Serialize(payload);
        using var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
        using var res = await http.PostAsync(endpoint, content, ct);
        if (!res.IsSuccessStatusCode)
        {
            var body = await res.Content.ReadAsStringAsync(ct);
            throw new Exception($"LLM endpoint error {res.StatusCode}: {body}");
        }
        var text = await res.Content.ReadAsStringAsync(ct);
        // 余計な説明行が混在する場合は先頭5行のみにトリム（空行除外）
        var lines = text.Replace("\r", "").Split('\n')
            .Select(l => l.Trim())
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Take(5)
            .ToList();
        return string.Join("\n", lines);
    }

    // Azure OpenAI Chat Completions 呼び出し
    private static async Task<string> GenerateTaglinesAzureAsync(string endpoint, string deployment, string apiKey, string apiVersion, string prompt, CancellationToken ct)
    {
        // endpoint 末尾スラッシュ除去
        endpoint = endpoint.TrimEnd('/');
        var url = $"{endpoint}/openai/deployments/{deployment}/chat/completions?api-version={apiVersion}";
        using var http = new HttpClient();
        http.DefaultRequestHeaders.Add("api-key", apiKey);
        http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        var bodyObj = new
        {
            messages = new object[]
            {
                new { role = "system", content = "あなたは日本語のプロフェッショナルコピーライターです。" },
                new { role = "user", content = prompt }
            },
            temperature = 0.8,
            top_p = 0.95,
            max_tokens = 400,
            n = 1
        };
        var json = System.Text.Json.JsonSerializer.Serialize(bodyObj);
        using var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
        using var res = await http.PostAsync(url, content, ct);
        var respText = await res.Content.ReadAsStringAsync(ct);
        if (!res.IsSuccessStatusCode)
        {
            throw new Exception($"Azure OpenAI error {res.StatusCode}: {respText}");
        }
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(respText);
            var root = doc.RootElement;
            var contentText = root.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? string.Empty;
            var lines = contentText.Replace("\r", "").Split('\n')
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Take(5)
                .ToList();
            return string.Join("\n", lines);
        }
        catch
        {
            // パース失敗時は素のテキストをライン整形
            var lines = respText.Replace("\r", "").Split('\n')
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Take(5)
                .ToList();
            return string.Join("\n", lines);
        }
    }
}

public record OneDriveUploadResult(bool IsSuccess, string? WebUrl, string? FileName, string? Error)
{
    public static OneDriveUploadResult CreateSuccess(string webUrl, string fileName) => new(true, webUrl, fileName, null);
    public static OneDriveUploadResult Fail(string error) => new(false, null, null, error);
}
