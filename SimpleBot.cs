using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.Extensions.DependencyInjection;
using PhraseXross.Services; // OneDriveExcelService
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;
using System;
using System.Text.RegularExpressions;
using Microsoft.Bot.Builder.Dialogs; // for DialogState
using System.IdentityModel.Tokens.Jwt; // JWT デコード用
using Microsoft.Bot.Builder.Integration.AspNet.Core; // CloudAdapter

public class SimpleBot : ActivityHandler
{
    private readonly Kernel? _kernel; // optional
    private readonly UserState? _userState; // Phase2: per-user state
    private readonly OneDriveExcelService? _oneDriveExcelService; // optional (OneDrive 連携)
        private readonly PhraseXross.Dialogs.MainDialog _mainDialog; // OAuthPrompt を含む Dialog
        private readonly ConversationState? _conversationState;
        private const string UnifiedWelcomeMessage = "こんにちは。今日はどんな言葉づくりをお手伝いしましょう？まずは、差し支えなければ『何の活動のためのコピーか』を教えてください。（例：イベント告知／販促キャンペーン／ブランド認知／採用 など）";

        public SimpleBot(Kernel? kernel = null, UserState? userState = null, OneDriveExcelService? oneDriveExcelService = null, PhraseXross.Dialogs.MainDialog? mainDialog = null, ConversationState? conversationState = null)
        {
            _kernel = kernel;
            _userState = userState;
            _oneDriveExcelService = oneDriveExcelService;
            _mainDialog = mainDialog ?? throw new ArgumentNullException(nameof(mainDialog));
            _conversationState = conversationState; // null 許容 (DI 差し込み漏れ対策)
        }

    private async Task<string?> TryGetDelegatedTokenAsync(ITurnContext turnContext, CancellationToken cancellationToken)
    {
        string? delegatedToken = null;
        var connectionName = Environment.GetEnvironmentVariable("BOT_OAUTH_CONNECTION_NAME") ?? "GraphDelegated";
        Console.WriteLine($"[DEBUG][SSO] TryGetDelegatedTokenAsync start userId='{turnContext.Activity.From?.Id}' channel='{turnContext.Activity.ChannelId}' conn='{connectionName}'");
        // 追加: メソッド探索や TurnState のキー一覧を詳細ログ（診断用）
        try
        {
            var turnStateTypes = turnContext.TurnState.Select(kv => kv.Value?.GetType()?.FullName ?? kv.Key ?? "<null>");
            Console.WriteLine("[DEBUG][SSO] TurnState types=" + string.Join(" | ", turnStateTypes));
        }
        catch { }
        ElicitationState? state = null;
        try
        {
            if (_userState != null)
            {
                var accessor = _userState.CreateProperty<ElicitationState>("ElicitationState");
                state = await accessor.GetAsync(turnContext, () => ElicitationState.CreateNew(), cancellationToken);
                state.LastTokenAttemptUtc = DateTimeOffset.UtcNow;
                state.LastTokenResult = null; // reset
                state.LastDelegatedTokenPreview = null;
            }
        }
        catch { /* state 取得失敗は致命ではない */ }
        try
        {
            object? userTokenClientObj = null;
            foreach (var kv in turnContext.TurnState)
            {
                var type = kv.Value?.GetType();
                if (type == null) continue;
                if (type.FullName == "Microsoft.Bot.Builder.Integration.AspNet.Core.UserTokenClient" || type.FullName == "Microsoft.Bot.Connector.Authentication.UserTokenClientImpl")
                {
                    userTokenClientObj = kv.Value; break;
                }
            }
            if (userTokenClientObj != null)
            {
                Console.WriteLine("[DEBUG][SSO] UserTokenClient found (reflection)");
                // パラメータ差異を考慮し複数候補を試す
                var methods = userTokenClientObj.GetType().GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                    .Where(m => m.Name == "GetUserTokenAsync")
                    .ToList();
                Console.WriteLine("[DEBUG][SSO] GetUserTokenAsync candidates paramSets=" + string.Join(" || ", methods.Select(m => string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name)))));
                foreach (var mi in methods)
                {
                    try
                    {
                        var ps = mi.GetParameters();
                        object?[] args;
                        switch (ps.Length)
                        {
                            case 5: // (userId, connectionName, channelId, magicCode, cancellationToken)
                                args = new object?[] { turnContext.Activity.From?.Id, connectionName, turnContext.Activity.ChannelId, null, cancellationToken }; break;
                            case 4: // (userId, connectionName, channelId, cancellationToken)
                                args = new object?[] { turnContext.Activity.From?.Id, connectionName, turnContext.Activity.ChannelId, cancellationToken }; break;
                            case 3: // (turnContext, connectionName, cancellationToken) など想定外 → スキップ
                                continue;
                            default:
                                continue;
                        }
                        var taskObj = mi.Invoke(userTokenClientObj, args);
                        if (taskObj is Task t)
                        {
                            await t.ConfigureAwait(false);
                            var resultProp = t.GetType().GetProperty("Result");
                            var tokenResponse = resultProp?.GetValue(t);
                            var tokenProp = tokenResponse?.GetType().GetProperty("Token");
                            delegatedToken = tokenProp?.GetValue(tokenResponse) as string;
                            if (!string.IsNullOrEmpty(delegatedToken))
                            {
                                Console.WriteLine($"[DEBUG][SSO] Token retrieved via signature paramCount={ps.Length}");
                                break;
                            }
                        }
                    }
                    catch (Exception exEach)
                    {
                        Console.WriteLine($"[DEBUG][SSO] GetUserTokenAsync variant failed: {exEach.GetType().Name}:{exEach.Message}");
                    }
                }
            }
            else
            {
                var adapter = turnContext.Adapter as BotAdapter;
                if (adapter != null)
                {
                    Console.WriteLine($"[DEBUG][SSO] Adapter reflection path adapterType='{adapter.GetType().FullName}'");
                    var mi = adapter.GetType().GetMethod("GetUserTokenAsync", new[] { typeof(ITurnContext), typeof(string), typeof(string), typeof(CancellationToken) });
                    if (mi != null)
                    {
                        var taskObj = mi.Invoke(adapter, new object?[] { turnContext, connectionName, null, cancellationToken });
                        if (taskObj is Task t2)
                        {
                            await t2.ConfigureAwait(false);
                            var resultProp = t2.GetType().GetProperty("Result");
                            var tokenResponse = resultProp?.GetValue(t2);
                            var tokenProp = tokenResponse?.GetType().GetProperty("Token");
                            delegatedToken = tokenProp?.GetValue(tokenResponse) as string;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("[DEBUG][SSO] Adapter is null or not BotAdapter");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[DEBUG][SSO] TryGetDelegatedTokenAsync failed: {ex.Message}");
            if (state != null) state.LastTokenResult = "exception:" + ex.GetType().Name;
        }
        if (!string.IsNullOrWhiteSpace(delegatedToken))
        {
            Console.WriteLine($"[DEBUG][SSO] Delegated token acquired (auto check). Length={delegatedToken.Length} Preview={MaskToken(delegatedToken)}");
            if (state != null && _userState != null)
            {
                state.LastTokenResult = "success";
                state.LastDelegatedTokenPreview = MaskToken(delegatedToken);
                try { await _userState.SaveChangesAsync(turnContext, false, cancellationToken); } catch { }
            }
        }
        else
        {
            if (state != null && _userState != null)
            {
                if (state.LastTokenResult == null) state.LastTokenResult = "null";
                try { await _userState.SaveChangesAsync(turnContext, false, cancellationToken); } catch { }
            }
        }
        return delegatedToken;
    }

    // 全ての初期ウェルカム送信を無効化する共通判定
    private static bool IsWelcomeDisabled()
    {
        var v = Environment.GetEnvironmentVariable("PX_DISABLE_WELCOME");
        // 未設定時はデフォルトで無効化 (true)
        if (string.IsNullOrWhiteSpace(v)) return true;
        return v.Equals("true", StringComparison.OrdinalIgnoreCase) || v == "1" || v.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<string?> TryGetSignInUrlAsync(ITurnContext turnContext, CancellationToken cancellationToken)
    {
        var connectionName = Environment.GetEnvironmentVariable("BOT_OAUTH_CONNECTION_NAME") ?? "GraphDelegated";
        try
        {
            object? userTokenClientObj = null;
            foreach (var kv in turnContext.TurnState)
            {
                var type = kv.Value?.GetType();
                if (type == null) continue;
                if (type.FullName == "Microsoft.Bot.Builder.Integration.AspNet.Core.UserTokenClient" || type.FullName == "Microsoft.Bot.Connector.Authentication.UserTokenClientImpl")
                {
                    userTokenClientObj = kv.Value; break;
                }
            }
            if (userTokenClientObj == null)
            {
                Console.WriteLine("[DEBUG][SSO] UserTokenClient not found for sign-in URL");
                return null;
            }
            // SDK バージョン差異によりパラメータが異なるため、すべての候補メソッドを試行
            var candidates = userTokenClientObj.GetType()
                .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                .Where(m => m.Name == "GetSignInResourceAsync")
                .ToList();
            if (candidates.Count == 0)
            {
                Console.WriteLine("[DEBUG][SSO] GetSignInResourceAsync not found via reflection");
                return null;
            }
            Console.WriteLine("[DEBUG][SSO] GetSignInResourceAsync candidates: " + string.Join(" | ", candidates.Select(c => string.Join(",", c.GetParameters().Select(p => p.ParameterType.Name)))));

            foreach (var mi in candidates)
            {
                try
                {
                    var ps = mi.GetParameters();
                    object?[] args;
                    // 代表的パターンを順次マッピング
                    switch (ps.Length)
                    {
                        case 5:
                            // (string connectionName, string userId, string channelId, string? finalRedirect, CancellationToken)
                            args = new object?[] { connectionName, turnContext.Activity.From?.Id, turnContext.Activity.ChannelId, null, cancellationToken };
                            break;
                        case 4:
                            // (string connectionName, string userId, string channelId, CancellationToken) もしくは finalRedirect 省略
                            args = new object?[] { connectionName, turnContext.Activity.From?.Id, turnContext.Activity.ChannelId, cancellationToken };
                            break;
                        case 3:
                            // (string connectionName, string userId, CancellationToken)
                            args = new object?[] { connectionName, turnContext.Activity.From?.Id, cancellationToken };
                            break;
                        default:
                            Console.WriteLine($"[DEBUG][SSO] Skip unsupported signature paramCount={ps.Length}");
                            continue;
                    }
                    var taskObj = mi.Invoke(userTokenClientObj, args);
                    if (taskObj is Task t)
                    {
                        await t.ConfigureAwait(false);
                        var resultProp = t.GetType().GetProperty("Result");
                        var signInResource = resultProp?.GetValue(t);
                        if (signInResource != null)
                        {
                            var linkProp = signInResource.GetType().GetProperty("SignInLink") ?? signInResource.GetType().GetProperty("Link");
                            var link = linkProp?.GetValue(signInResource) as string;
                            if (!string.IsNullOrWhiteSpace(link))
                            {
                                Console.WriteLine($"[DEBUG][SSO] SignInLink obtained via signature paramCount={ps.Length} length={link.Length}");
                                return link;
                            }
                        }
                    }
                }
                catch (Exception exEach)
                {
                    Console.WriteLine($"[DEBUG][SSO] GetSignInResourceAsync candidate failed: {exEach.GetType().Name}:{exEach.Message}");
                    // 次の候補を試す
                }
            }
            Console.WriteLine("[DEBUG][SSO] All GetSignInResourceAsync attempts failed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[DEBUG][SSO] TryGetSignInUrlAsync failed: {ex.Message}");
        }
        return null;
    }

    private static string MaskToken(string token)
    {
        if (string.IsNullOrEmpty(token)) return "";
        if (token.Length <= 8) return new string('*', token.Length);
        return token.Substring(0, 4) + "..." + token.Substring(token.Length - 4);
    }

    private static string TrimForDump(string text, int max)
    {
        if (string.IsNullOrEmpty(text)) return "";
        if (text.Length <= max) return text;
        return text.Substring(0, max) + "...";
    }

    private static string MaskMid(string value)
    {
        if (string.IsNullOrEmpty(value) || value == "<null>") return value ?? "";
        if (value.Length <= 6) return new string('*', value.Length);
        return value.Substring(0,3) + "***" + value.Substring(value.Length-3);
    }

    protected override async Task OnEventActivityAsync(ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
    {
        Console.WriteLine($"[DEBUG][SSO][Event] name='{turnContext.Activity.Name}' channel='{turnContext.Activity.ChannelId}' type='{turnContext.Activity.Type}'");
        // tokens/response など OAuthPrompt が処理すべきイベントを Dialog に流す
        if (_conversationState != null && _mainDialog != null &&
            (string.Equals(turnContext.Activity.Name, "tokens/response", StringComparison.OrdinalIgnoreCase)))
        {
            var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
            Console.WriteLine("[OAUTH][EventPipeline] Forwarding tokens/response to Dialog");
            await _mainDialog.RunAsync(turnContext, dialogStateAccessor, cancellationToken);
            await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
        }

        // トークン取得後の Excel 再開など
        if (_userState != null && _oneDriveExcelService != null)
        {
            var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
            var state = await accessor.GetAsync(turnContext, () => ElicitationState.CreateNew(), cancellationToken);
            if (state.WaitingForSignIn && !state.Step8Completed)
            {
                // OAuthPrompt により state.DelegatedGraphToken が埋まっていれば再開
                if (!string.IsNullOrWhiteSpace(state.DelegatedGraphToken))
                {
                    var token = state.DelegatedGraphToken;
                    state.WaitingForSignIn = false;                    
                    await turnContext.SendActivityAsync(MessageFactory.Text("サインインが完了しました。Excel 出力を再開します。"), cancellationToken);
                    var result = await _oneDriveExcelService.CreateAndFillExcelAsync(null, state.TaglineSummaryJson, cancellationToken, token);
                    if (result.IsSuccess && !string.IsNullOrWhiteSpace(result.WebUrl))
                    {
                        var done = $"Excel出力完了: {result.WebUrl}";
                        await turnContext.SendActivityAsync(MessageFactory.Text(done, done), cancellationToken);
                        state.Step8Completed = true;
                        state.CompletedUtc = DateTimeOffset.UtcNow;
                        state.History.Clear();
                        var guidance = "このセッションは完了しました。新しい案件を開始します。キャッチコピー作成の目的を一言で教えてください。";
                        var newState = ElicitationState.CreateNew();
                        await accessor.SetAsync(turnContext, newState, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                        await turnContext.SendActivityAsync(MessageFactory.Text(guidance, guidance), cancellationToken);
                        return;
                    }
                    else
                    {
                        var err = $"Excel出力失敗: {result.Error ?? "不明なエラー"}";
                        await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
                    }
                }
            }
            await accessor.SetAsync(turnContext, state, cancellationToken);
            await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
        }
        await base.OnEventActivityAsync(turnContext, cancellationToken);
    }

    protected override async Task<InvokeResponse> OnInvokeActivityAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
    {
        // signin/tokenExchange 等にも対応
        Console.WriteLine($"[DEBUG][SSO][Invoke] name='{turnContext.Activity.Name}' channel='{turnContext.Activity.ChannelId}' type='{turnContext.Activity.Type}'");
        var isSignInInvoke = string.Equals(turnContext.Activity.Name, "signin/tokenExchange", StringComparison.OrdinalIgnoreCase) || string.Equals(turnContext.Activity.Name, "signin/verifyState", StringComparison.OrdinalIgnoreCase);
        if (isSignInInvoke && _conversationState != null && _mainDialog != null)
        {
            var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
            Console.WriteLine("[OAUTH][InvokePipeline] Forwarding invoke to Dialog (OAuthPrompt)");
            await _mainDialog.RunAsync(turnContext, dialogStateAccessor, cancellationToken);
            await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
        }

        if (isSignInInvoke && _userState != null && _oneDriveExcelService != null)
        {
            var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
            var state = await accessor.GetAsync(turnContext, () => ElicitationState.CreateNew(), cancellationToken);
            if (state.WaitingForSignIn && !state.Step8Completed && !string.IsNullOrWhiteSpace(state.DelegatedGraphToken))
            {
                var token = state.DelegatedGraphToken;
                state.WaitingForSignIn = false;
                var result = await _oneDriveExcelService.CreateAndFillExcelAsync(null, state.TaglineSummaryJson, cancellationToken, token);
                if (result.IsSuccess && !string.IsNullOrWhiteSpace(result.WebUrl))
                {
                    var done = $"Excel出力完了: {result.WebUrl}";
                    await turnContext.SendActivityAsync(MessageFactory.Text(done, done), cancellationToken);
                    state.Step8Completed = true;
                    state.CompletedUtc = DateTimeOffset.UtcNow;
                    state.History.Clear();
                    var guidance = "このセッションは完了しました。新しい案件を開始します。キャッチコピー作成の目的を一言で教えてください。";
                    var newState = ElicitationState.CreateNew();
                    await accessor.SetAsync(turnContext, newState, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    await turnContext.SendActivityAsync(MessageFactory.Text(guidance, guidance), cancellationToken);
                    return new InvokeResponse { Status = 200 };
                }
                else
                {
                    var err = $"Excel出力失敗: {result.Error ?? "不明なエラー"}";
                    await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
                }
                await accessor.SetAsync(turnContext, state, cancellationToken);
                await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
            }
        }
        var resp = await base.OnInvokeActivityAsync(turnContext, cancellationToken);
        return resp ?? new InvokeResponse { Status = 200 };
    }

    protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        var text = turnContext.Activity.Text ?? string.Empty;
        Console.WriteLine($"[USER_INPUT] {text}");

        if (string.IsNullOrWhiteSpace(text))
        {
            var emptyMessage = "入力が空です。もう一度入力してください。";
            Console.WriteLine($"[USER_MESSAGE] {emptyMessage}");
            await turnContext.SendActivityAsync(MessageFactory.Text(emptyMessage, emptyMessage), cancellationToken);
            return;
        }

        if (_kernel == null)
        {
            // ここに到達するのは設定漏れ。ユーザー通知はせずログのみ（本番では起動時に弾く想定）。
            Console.WriteLine("[FATAL] Kernel not registered. Configure Semantic Kernel before starting the bot.");
            return; // 何も返さず黙る（開発時にログで気付く）
        }

        try
        {
            // 会話状態取得（ユーザー/アシスタントの履歴を保持）
            var state = new ElicitationState();
            if (_userState != null)
            {
                var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                state = await accessor.GetAsync(turnContext, () => ElicitationState.CreateNew(), cancellationToken);
            }

            // --- 手法A: 会話開始フォールバック初期ガイダンス ---
            // Teams 等で conversationUpdate (membersAdded) が届かず Welcome が未送信のまま
            // ユーザーが最初の発話をしてきた場合に、ここで静的な Step1 導入メッセージを一度だけ挿入する。
            // 条件:
            //   - まだどの Step も完了していない
            //   - WelcomeSent == false
            //   - 過去履歴/質問カウントが 0（= まだボットからのヒアリング未開始）
            //   - 入力がコマンド (/ で始まる) ではない
            //   - Reset 系での再開直後を除く (History.Count == 0 を条件とする)
            var isAnyStepCompleted = state.Step1Completed || state.Step2Completed || state.Step3Completed || state.Step4Completed || state.Step5Completed || state.Step6Completed || state.Step7Completed || state.Step8Completed;
            var disableFallbackWelcome = IsWelcomeDisabled() || string.Equals(Environment.GetEnvironmentVariable("PX_DISABLE_FALLBACK_WELCOME"), "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(Environment.GetEnvironmentVariable("PX_DISABLE_FALLBACK_WELCOME"), "1", StringComparison.OrdinalIgnoreCase);
            if (!disableFallbackWelcome && !isAnyStepCompleted && !state.WelcomeSent && state.History.Count == 0 && state.Step1QuestionCount == 0 && !string.IsNullOrWhiteSpace(text) && !text.TrimStart().StartsWith("/"))
            {
                var fallbackWelcome = UnifiedWelcomeMessage;
                // ユーザー最初の発話が既に同内容（コピペ等）の場合は抑止
                var normalizedUser = text.Replace("\\s+", " ").Trim();
                var normalizedWelcome = fallbackWelcome.Replace("\\s+", " ").Trim();
                if (!normalizedUser.Equals(normalizedWelcome, StringComparison.Ordinal))
                {
                    Console.WriteLine($"[USER_MESSAGE][FallbackWelcome] {fallbackWelcome}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(fallbackWelcome, fallbackWelcome), cancellationToken);
                    state.History.Add($"Assistant: {fallbackWelcome}"); // 重複抑制のため履歴へも記録
                    TrimHistory(state.History, 30);
                }
                state.WelcomeSent = true; // 二重送信防止
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                // 続けてユーザー発話を Step1 の回答として既存ロジックに流す（return しない）
            }

            // 先に /forceauth を処理（auto-run と競合させない）
            var rawLower = text.Trim().ToLowerInvariant();
            if (rawLower == "/forceauth")
            {
                if (_conversationState != null && _mainDialog != null)
                {
                    var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
                    Console.WriteLine("[OAUTH][Force] MainDialog invoked by user (pre-auto-run phase)");
                    await _mainDialog.RunAsync(turnContext, dialogStateAccessor, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    await turnContext.SendActivityAsync(MessageFactory.Text("サインイン処理を開始しました。カードが表示されない場合は Teams のキャッシュクリア（Ctrl+R または 完全再起動）を試してください。", "サインイン処理を開始しました。"), cancellationToken);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text("会話状態が初期化されていないため OAuthPrompt を開始できません。", "会話状態が初期化されていないため OAuthPrompt を開始できません。"), cancellationToken);
                }
                return;
            }

            // /signin コマンド: Step 進行状況に関係なく明示的に OAuthPrompt を起動
            if (rawLower == "/signin")
            {
                if (_conversationState != null && _mainDialog != null)
                {
                    var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
                    Console.WriteLine("[OAUTH][SigninCmd] MainDialog invoked by user (/signin)");
                    await _mainDialog.RunAsync(turnContext, dialogStateAccessor, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    await turnContext.SendActivityAsync(MessageFactory.Text("サインインカードを表示します。表示されない場合は /oauthdiag で状態を確認してください。", "サインインカードを表示します。"), cancellationToken);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text("会話状態が初期化されていないため OAuthPrompt を開始できません。", "会話状態が初期化されていないため OAuthPrompt を開始できません。"), cancellationToken);
                }
                return;
            }

            // === OAuthPrompt 実行トリガ ===
            // 条件: まだトークン未取得 かつ (Step6 以降に到達 / サインイン待機フラグ) の場合に自動実行。
            if (_conversationState != null && _mainDialog != null)
            {
                // /forceauth コマンドのターンでは自動起動をスキップ（同一ターン二重実行防止）
                var isForceAuthCommand = string.Equals(text.Trim(), "/forceauth", StringComparison.OrdinalIgnoreCase);
                var needAuth = !isForceAuthCommand && string.IsNullOrWhiteSpace(state.DelegatedGraphToken) && (state.Step6Completed || state.Step7Completed || state.WaitingForSignIn);
                if (needAuth)
                {
                    var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
                    Console.WriteLine("[OAUTH][AutoRun] MainDialog invoked (reason=needAuth)");
                    await _mainDialog.RunAsync(turnContext, dialogStateAccessor, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
            }

            // === セッション完了後の自動リセットガード ===
            // 直前に Step8 が完了しているが、新しいガイダンス前にユーザーが素早く入力した場合、旧 state の履歴を引きずる可能性がある。
            // Completed セッションにユーザーが通常メッセージを送ったら自動で新規セッションを開始し、その最初のメッセージとして扱う。
            if (state.Step8Completed && state.CompletedUtc != null && (state.Step1Completed || state.Step2Completed || state.Step3Completed || state.Step4Completed || state.Step5Completed || state.Step6Completed || state.Step7Completed))
            {
                var oldSessionId = state.SessionId;
                state = ElicitationState.CreateNew();
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                Console.WriteLine($"[AUTO_RESET] Prior session {oldSessionId} was completed. Started new session {state.SessionId} before processing user input.");
            }

            // /reset コマンド（大小文字・全角半角空白を許容簡易）
            var trimmed = text.Trim();
            var lower = trimmed.ToLowerInvariant();
            if (lower == "/authstatus")
            {
                var diag = new List<string>();
                diag.Add("— Auth Status —");
                diag.Add($"WaitingForSignIn: {state.WaitingForSignIn}");
                diag.Add($"ConnectionName: {Environment.GetEnvironmentVariable("BOT_OAUTH_CONNECTION_NAME") ?? "GraphDelegated"}");
                diag.Add($"Step8Completed: {state.Step8Completed}");
                diag.Add($"LastTokenAttemptUtc: {state.LastTokenAttemptUtc:O}");
                    if (!string.IsNullOrWhiteSpace(state.DelegatedGraphToken)) diag.Add("DelegatedGraphToken: (present)");
                if (!string.IsNullOrWhiteSpace(state.LastTokenResult)) diag.Add($"LastTokenResult: {state.LastTokenResult}");
                if (!string.IsNullOrWhiteSpace(state.LastDelegatedTokenPreview)) diag.Add($"TokenPreview: {state.LastDelegatedTokenPreview}");
                var msgDiag = string.Join("\n", diag);
                Console.WriteLine("[DEBUG][SSO] /authstatus -> " + msgDiag.Replace("\n", " | "));
                await turnContext.SendActivityAsync(MessageFactory.Text(msgDiag, msgDiag), cancellationToken);
                return;
            }

            if (lower == "/oauthdiag")
            {
                var lines = new List<string>();
                lines.Add("— OAuth Diagnostics —");
                lines.Add($"ConnectionName: {Environment.GetEnvironmentVariable("BOT_OAUTH_CONNECTION_NAME") ?? "GraphDelegated"}");
                lines.Add($"PromptStartCount: {state.OAuthPromptStartCount}");
                lines.Add($"PromptLastAttemptUtc: {state.OAuthPromptLastAttemptUtc:O}");
                lines.Add($"LastTokenAcquiredUtc: {state.LastTokenAcquiredUtc:O}");
                lines.Add($"DelegatedTokenPresent: {!string.IsNullOrWhiteSpace(state.DelegatedGraphToken)}");
                lines.Add($"WaitingForSignIn: {state.WaitingForSignIn}");
                if (state.DelegatedTokenExpiresUtc != null)
                {
                    var remain = state.DelegatedTokenExpiresUtc.Value - DateTimeOffset.UtcNow;
                    if (remain < TimeSpan.Zero) remain = TimeSpan.Zero;
                    lines.Add($"TokenExpiresUtc: {state.DelegatedTokenExpiresUtc:O}");
                    lines.Add($"TokenRemaining: {remain:hh\\:mm\\:ss}");
                }
                var txt = string.Join("\n", lines);
                await turnContext.SendActivityAsync(MessageFactory.Text(txt, txt), cancellationToken);
                return;
            }

            if (lower == "/expiretoken")
            {
                state.DelegatedTokenExpiresUtc = DateTimeOffset.UtcNow.AddSeconds(-10);
                var msg = "ローカルでトークン期限を強制的に失効状態にしました。次回 Step8 実行で再取得動作を確認できます。";
                await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }

            if (lower == "/droptoken")
            {
                state.DelegatedGraphToken = null;
                state.DelegatedTokenExpiresUtc = null;
                var msg = "保存していたトークンを破棄しました。次の Step8 でサインインカードが表示されます。";
                await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }

            if (lower == "/forceauth")
            {
                if (_conversationState != null && _mainDialog != null)
                {
                    var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
                    Console.WriteLine("[OAUTH][Force] MainDialog invoked by user");
                    await _mainDialog.RunAsync(turnContext, dialogStateAccessor, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    await turnContext.SendActivityAsync(MessageFactory.Text("サインイン処理を開始しました。カードが表示されない場合は Teams のキャッシュクリアやアプリ再インストールをお試しください。", "サインイン処理を開始しました。"), cancellationToken);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text("会話状態が初期化されていないため OAuthPrompt を開始できません。", "会話状態が初期化されていないため OAuthPrompt を開始できません。"), cancellationToken);
                }
                return;
            }

            if (lower == "/signin")
            {
                // /signin コマンドは OAuthPrompt へ移行したため廃止
            }

            if (lower == "/tokenretry")
            {
                // /tokenretry も削除（OAuthPrompt へ統一）
            }

            if (lower == "/signinurl")
            {
                // /signinurl も削除
            }

            // 手動で Step8 (Excel 出力) を試行するためのデバッグコマンド
            if (lower == "/excel" || lower == "/step8")
            {
                // /excel, /step8 コマンドは削除 (自動フローのみ)
            }

            // SSO / OAuth の設定ガイドを表示
            if (lower == "/ssohelp" || lower == "/oauthhelp")
            {
                var lines = new List<string>();
                lines.Add("— SSO / OAuth 診断ガイド —");
                lines.Add("1. Azure Portal > Bot Channels Registration で OAuth Connection (例: GraphDelegated) を作成");
                lines.Add("   - Service Provider: Azure Active Directory v2");
                lines.Add("   - Scopes 例: offline_access openid profile User.Read Files.ReadWrite.All Sites.ReadWrite.All");
                lines.Add("2. BOT_OAUTH_CONNECTION_NAME 環境変数が接続名と一致しているか確認");
                lines.Add("3. Teams アプリの再インストール / キャッシュクリア (サインインカードが出ない場合)");
                lines.Add("4. /configcheck で UserTokenClientInTurnState=True を確認 (False の場合は AppId/Password 不整合)");
                lines.Add("5. /signin コマンドで明示的にサインインカードを表示できます");
                lines.Add("6. サインイン後、自動イベント (signin/verifyState, tokens/response) が届くと Step8 再開");
                lines.Add("7. 失敗時 /tokenretry で再取得、/signinurl でブラウザ直接サインインを試行");
                lines.Add("8. AAD 側リダイレクト URL: https://token.botframework.com/.auth/web/redirect を含めること");
                var help = string.Join("\n", lines);
                await turnContext.SendActivityAsync(MessageFactory.Text(help, help), cancellationToken);
                return;
            }

            if (lower == "/configcheck")
            {
                string GetEnv(string k) => Environment.GetEnvironmentVariable(k) ?? "<null>";
                var lines = new List<string>();
                lines.Add("— Config Check —");
                lines.Add($"MicrosoftAppId: {MaskMid(GetEnv("MicrosoftAppId"))}");
                var pwd = GetEnv("MicrosoftAppPassword");
                lines.Add($"MicrosoftAppPassword: {(pwd=="<null>"?"<null>":$"len={pwd.Length}")}");
                lines.Add($"MicrosoftAppTenantId: {GetEnv("MicrosoftAppTenantId")}");
                lines.Add($"MicrosoftAppType: {GetEnv("MicrosoftAppType")}");
                lines.Add($"BOT_OAUTH_CONNECTION_NAME: {GetEnv("BOT_OAUTH_CONNECTION_NAME")}");
                // TurnState introspection
                bool tokenClientFound = false;
                foreach (var kv in turnContext.TurnState)
                {
                    var t = kv.Value?.GetType();
                    if (t?.FullName == "Microsoft.Bot.Builder.Integration.AspNet.Core.UserTokenClient" || t?.FullName == "Microsoft.Bot.Connector.Authentication.UserTokenClientImpl") tokenClientFound = true;
                }
                lines.Add($"UserTokenClientInTurnState: {tokenClientFound}");
                if (!tokenClientFound)
                {
                    lines.Add("(原因候補) AppId/Password 未設定 または 認証無効モード / OAuth Connection 不整合");
                }
                var txt = string.Join("\n", lines);
                Console.WriteLine("[DEBUG][CFG] /configcheck -> " + txt.Replace("\n"," | "));
                await turnContext.SendActivityAsync(MessageFactory.Text(txt, txt), cancellationToken);
                return;
            }

            if (lower == "/turnstate")
            {
                var keys = new List<string>();
                foreach (var kv in turnContext.TurnState)
                {
                    var t = kv.Value?.GetType();
                    keys.Add(t?.FullName ?? kv.Key ?? "<unknown>");
                }
                if (keys.Count == 0) keys.Add("<empty>");
                var txt = "— TurnState Keys —\n" + string.Join("\n", keys);
                Console.WriteLine("[DEBUG][TURNSTATE] " + txt.Replace("\n"," | "));
                await turnContext.SendActivityAsync(MessageFactory.Text(txt, txt), cancellationToken);
                return;
            }

            if (lower == "/authdump" || lower == "/activity")
            {
                var a = turnContext.Activity;
                var lines = new List<string>();
                lines.Add("— Activity Dump —");
                lines.Add($"Type: {a.Type}");
                if (a is IEventActivity ev && !string.IsNullOrWhiteSpace(ev.Name)) lines.Add($"Name: {ev.Name}");
                else if (a is IInvokeActivity inv && !string.IsNullOrWhiteSpace(inv.Name)) lines.Add($"Name: {inv.Name}");
                lines.Add($"ChannelId: {a.ChannelId}");
                lines.Add($"From.Id: {a.From?.Id}");
                lines.Add($"Conversation.Id: {a.Conversation?.Id}");
                if (!string.IsNullOrWhiteSpace(a.Text)) lines.Add($"Text: {TrimForDump(a.Text,120)}");
                if (a.Value != null) lines.Add($"ValueType: {a.Value.GetType().Name}");
                if (a.Attachments != null && a.Attachments.Any()) lines.Add($"Attachments: {a.Attachments.Count}");
                var dump = string.Join("\n", lines);
                Console.WriteLine("[DEBUG][DUMP] " + dump.Replace("\n"," | "));
                await turnContext.SendActivityAsync(MessageFactory.Text(dump, dump), cancellationToken);
                return;
            }

            // リセットトリガー
            // - 『リセット』『最初から』は“全文字一致のみ”で発火（ユーザー要望）
            // - 互換維持のため、従来の英語/スラッシュ系キーワードも残す
            //   ただし注意: Teams クライアントには /reset（プレゼンス用）のビルトインがあり、
            //   そちらが優先されるとボットにメッセージが届かない場合があります。
            //   それでも他チャネル/クライアントや将来の環境を考慮して残置します。
            if (text == "リセット" || text == "最初から"
                || lower == "/reset" || lower == "reset"
                || lower == "新規" || lower == "/new" || lower == "new"
                || lower == "/start" || lower == "start"
                || lower == "別件" || lower == "別の")
            {
                // 既存セッションを完了扱い (未完了なら Cancelled コメント相当)
                if (!state.Step8Completed && state.CompletedUtc == null)
                {
                    state.CompletedUtc = DateTimeOffset.UtcNow; // 未完了終了
                }
                // 新規セッション再生成
                state = ElicitationState.CreateNew();
                var msg = "セッションをリセットしました。まずキャッチコピー作成の目的を一言で教えてください。";
                Console.WriteLine($"[USER_MESSAGE] {msg}");
                await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }

            // /hardreset は履歴を強制クリア（セッション完了扱い + 新規）
            if (lower == "/hardreset" || lower == "hardreset")
            {
                state.CompletedUtc = DateTimeOffset.UtcNow;
                state = ElicitationState.CreateNew();
                state.History.Clear();
                var msg = "履歴を完全クリアしました。改めて目的を一言で教えてください。";
                Console.WriteLine($"[USER_MESSAGE] {msg}");
                await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }


            // 新規セッション直後（全ステップ未完）で履歴が残っている場合は防御的にクリア（想定外な残留対策）
            if (!state.Step1Completed && !state.Step2Completed && !state.Step3Completed && !state.Step4Completed && !state.Step5Completed && !state.Step6Completed && !state.Step7Completed && !state.Step8Completed && state.History.Count > 0)
            {
                Console.WriteLine("[DEFENSIVE] Unexpected residual history detected on fresh session. Clearing.");
                state.History.Clear();
            }

            // 履歴にユーザー発話を追加
            state.History.Add($"User: {text}");
            TrimHistory(state.History, 30);
            var transcript = string.Join("\n", state.History);

            // 既にStep1が完了している場合はStep2（ターゲット）フローに分岐
            if (state.Step1Completed && !state.Step2Completed)
            {
                // まず、ターゲットが十分に固まっているかを評価
                var tEval = await EvaluateTargetAsync(_kernel, transcript, state.Step1SummaryJson, cancellationToken);
                if (tEval != null) state.LastTargetReason = tEval.Reason;
                if (tEval != null && tEval.IsSatisfied)
                {
                    state.FinalTarget = string.IsNullOrWhiteSpace(tEval.Target) ? state.FinalTarget : tEval.Target;

                    // 簡潔に承知のみ（Step2はAI生成の承知文は使わず最小限の固定文）
                    var tAck = $"ターゲット像、承知しました。ありがとうございます。\n- {state.FinalTarget}";
                    Console.WriteLine($"[USER_MESSAGE] {tAck}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(tAck, tAck), cancellationToken);

                    state.Step2Completed = true; // ターゲット完了

                    var consolidatedAfterStep2 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep2}");
                    await SendSummaryAsync(turnContext, consolidatedAfterStep2, cancellationToken);

                    // Step3（媒体/利用シーン）へ最初の短い質問を投げる（元の自動遷移ロジックを復元）
                    var mFirst = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, state.LastMediaReason, cancellationToken);
                    if (string.IsNullOrWhiteSpace(mFirst))
                    {
                        mFirst = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, state.LastMediaReason, cancellationToken);
                    }
                    if (!string.IsNullOrWhiteSpace(mFirst))
                    {
                        state.History.Add($"Assistant: {mFirst}");
                        TrimHistory(state.History, 30);
                        state.Step3QuestionCount++;

                        Console.WriteLine($"[USER_MESSAGE] {mFirst}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(mFirst, mFirst), cancellationToken);
                    }
                    // 状態保存
                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 既出の手がかりを踏まえて次のターゲット質問を生成
                var tAsk = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, state.LastTargetReason, cancellationToken);

                if (string.IsNullOrWhiteSpace(tAsk))
                {
                    // 1回だけAIに再試行
                    tAsk = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, state.LastTargetReason, cancellationToken);
                }

                if (!string.IsNullOrWhiteSpace(tAsk))
                {
                    state.History.Add($"Assistant: {tAsk}");
                    TrimHistory(state.History, 30);
                    state.Step2QuestionCount++;

                    // 状態保存
                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }

                    Console.WriteLine($"[USER_MESSAGE] {tAsk}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(tAsk, tAsk), cancellationToken);
                }
                return;
            }

            // Step3（媒体/利用シーン）フロー
            if (state.Step1Completed && state.Step2Completed && !state.Step3Completed)
            {
                // 既出の usageContext などを踏まえて十分性を評価
                var mEval = await EvaluateMediaAsync(_kernel, transcript, state.Step1SummaryJson, cancellationToken);
                if (mEval != null) state.LastMediaReason = mEval.Reason;
                if (mEval != null && mEval.IsSatisfied)
                {
                    state.FinalUsageContext = string.IsNullOrWhiteSpace(mEval.MediaOrContext) ? state.FinalUsageContext : mEval.MediaOrContext;

                    var mAck = $"媒体／利用シーン、承知しました。ありがとうございます。\n- {state.FinalUsageContext}";
                    Console.WriteLine($"[USER_MESSAGE] {mAck}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(mAck, mAck), cancellationToken);

                    state.Step3Completed = true; // 媒体完了

                    var consolidatedAfterStep3 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep3}");
                    await SendSummaryAsync(turnContext, consolidatedAfterStep3, cancellationToken);

                    // Step4（コア価値）初回質問を自動生成して即時遷移（Step2->Step3 と対称性を保つ）
                    try
                    {
                        var coreFirst = await GenerateNextCoreQuestionAsync(
                            _kernel,
                            transcript,
                            state.Step1SummaryJson,
                            state.FinalTarget,
                            state.FinalUsageContext,
                            state.Step4QuestionCount,
                            null,
                            cancellationToken);
                        if (string.IsNullOrWhiteSpace(coreFirst))
                        {
                            // 一度だけ再試行
                            coreFirst = await GenerateNextCoreQuestionAsync(
                                _kernel,
                                transcript,
                                state.Step1SummaryJson,
                                state.FinalTarget,
                                state.FinalUsageContext,
                                state.Step4QuestionCount,
                                null,
                                cancellationToken);
                        }
                        if (IsInvalidCoreQuestion(coreFirst)) coreFirst = string.Empty;
                        if (string.IsNullOrWhiteSpace(coreFirst))
                        {
                            // 再試行1: 明確化指示
                            var booster1 = new ChatHistory();
                            booster1.AddSystemMessage("前回の出力が空または無効でした。提供価値を一言で補足してもらうための短い質問を1つだけ返してください。30文字以内。列挙禁止。");
                            booster1.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
                            var b1 = await InvokeAndLogAsync(_kernel, booster1, cancellationToken, "S4/CORE:Q-INIT-BOOST1");
                            coreFirst = b1?.Content?.Trim();
                            if (IsInvalidCoreQuestion(coreFirst)) coreFirst = string.Empty;
                        }
                        if (string.IsNullOrWhiteSpace(coreFirst))
                        {
                            // 再試行2: さらに制約強調
                            var booster2 = new ChatHistory();
                            booster2.AddSystemMessage("再試行: コアとなる価値を確定するための焦点質問を 1 つ返してください。『どれが近いですか』などの列挙は避け、単一質問のみ。");
                            booster2.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
                            var b2 = await InvokeAndLogAsync(_kernel, booster2, cancellationToken, "S4/CORE:Q-INIT-BOOST2");
                            coreFirst = b2?.Content?.Trim();
                            if (IsInvalidCoreQuestion(coreFirst)) coreFirst = string.Empty;
                        }
                        if (string.IsNullOrWhiteSpace(coreFirst))
                        {
                            coreFirst = "提供価値の核心は？"; // 最終フェイルセーフ（短く中立）
                            Console.WriteLine("[DEBUG] Core initial question ultimate fallback used (all retries empty)");
                        }
                        state.History.Add($"Assistant: {coreFirst}");
                        TrimHistory(state.History, 30);
                        state.Step4QuestionCount++;
                        Console.WriteLine($"[USER_MESSAGE] {coreFirst}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(coreFirst, coreFirst), cancellationToken);
                    }
                    catch (Exception exCore)
                    {
                        var fb = "次に『提供価値・差別化のコア』を一言で教えてください。例：地域ならではの雰囲気／初心者歓迎の安心感 等。";
                        Console.WriteLine($"[WARN] Core question generation failed: {exCore.Message}");
                        Console.WriteLine($"[USER_MESSAGE] {fb}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(fb, fb), cancellationToken);
                        // 失敗時も前進させるためにカウントだけ進める
                        state.Step4QuestionCount++;
                    }

                    // 状態保存
                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 既出の usageContext を活かして次の質問を生成
                var mAsk = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, state.LastMediaReason, cancellationToken);
                if (string.IsNullOrWhiteSpace(mAsk))
                {
                    mAsk = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, state.LastMediaReason, cancellationToken);
                }
                if (!string.IsNullOrWhiteSpace(mAsk))
                {
                    state.History.Add($"Assistant: {mAsk}");
                    TrimHistory(state.History, 30);
                    state.Step3QuestionCount++;

                    // 状態保存
                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }

                    Console.WriteLine($"[USER_MESSAGE] {mAsk}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(mAsk, mAsk), cancellationToken);
                }
                return;
            }

            // Step4（提供価値・差別化のコア）フロー
            if (state.Step1Completed && state.Step2Completed && state.Step3Completed && !state.Step4Completed)
            {
                var cEval = await EvaluateCoreAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, cancellationToken);
                if (cEval != null) state.LastCoreReason = cEval.Reason;
                if (cEval != null && cEval.IsSatisfied)
                {
                    state.FinalCoreValue = string.IsNullOrWhiteSpace(cEval.Core) ? state.FinalCoreValue : cEval.Core;

                    var cAck = $"コア（提供価値・差別化）、承知しました。ありがとうございます。\n- {state.FinalCoreValue}";
                    Console.WriteLine($"[USER_MESSAGE] {cAck}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(cAck, cAck), cancellationToken);

                    state.Step4Completed = true; // Step4完了

                    var consolidatedAfterStep4 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep4}");
                    await SendSummaryAsync(turnContext, consolidatedAfterStep4, cancellationToken);

                    // Step5（制約事項）初回: ハードコードせず LLM 生成を直接利用（ユーザー要望により固定文撤去）
                    if (!state.Step5Completed)
                    {
                        var firstConstraintQ = await GenerateNextConstraintsQuestionAsync(
                            _kernel,
                            transcript,
                            state.Step1SummaryJson,
                            state.FinalTarget,
                            state.FinalUsageContext,
                            state.FinalCoreValue,
                            state.Step5QuestionCount,
                            state.LastConstraintsReason,
                            cancellationToken);
                        if (string.IsNullOrWhiteSpace(firstConstraintQ))
                        {
                            // モデル失敗時の最低限フォールバック（誘導語『特にない』等は含めない）
                            firstConstraintQ = "制約事項（文字数・避けたい表現・法規・文化配慮など）で既に決まっている点があれば教えてください。";
                        }
                        state.History.Add($"Assistant: {firstConstraintQ}");
                        TrimHistory(state.History, 30);
                        state.Step5QuestionCount++;
                        Console.WriteLine($"[USER_MESSAGE] {firstConstraintQ}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(firstConstraintQ, firstConstraintQ), cancellationToken);
                    }

                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 既出の手がかりを活かして次の質問を生成
                var cAsk = await GenerateNextCoreQuestionAsync(
                    _kernel,
                    transcript,
                    state.Step1SummaryJson,
                    state.FinalTarget,
                    state.FinalUsageContext,
                    state.Step4QuestionCount,
                    cEval?.Reason,
                    cancellationToken);

                if (string.IsNullOrWhiteSpace(cAsk))
                {
                    cAsk = await GenerateNextCoreQuestionAsync(
                        _kernel,
                        transcript,
                        state.Step1SummaryJson,
                        state.FinalTarget,
                        state.FinalUsageContext,
                        state.Step4QuestionCount,
                        cEval?.Reason,
                        cancellationToken);
                }

                if (IsInvalidCoreQuestion(cAsk)) cAsk = string.Empty;
                if (string.IsNullOrWhiteSpace(cAsk))
                {
                    var booster1 = new ChatHistory();
                    booster1.AddSystemMessage("前回のコア質問が空/無効でした。30文字以内で焦点を一つに絞った質問を1つ返してください。");
                    booster1.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
                    var b1 = await InvokeAndLogAsync(_kernel, booster1, cancellationToken, "S4/CORE:Q-BOOST1B");
                    cAsk = b1?.Content?.Trim();
                    if (IsInvalidCoreQuestion(cAsk)) cAsk = string.Empty;
                }
                if (string.IsNullOrWhiteSpace(cAsk))
                {
                    var booster2 = new ChatHistory();
                    booster2.AddSystemMessage("再試行: コア提供価値を確定するための一点集中の質問を1つだけ返してください。列挙や複数質問は禁止。");
                    booster2.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
                    var b2 = await InvokeAndLogAsync(_kernel, booster2, cancellationToken, "S4/CORE:Q-BOOST2B");
                    cAsk = b2?.Content?.Trim();
                    if (IsInvalidCoreQuestion(cAsk)) cAsk = string.Empty;
                }
                if (string.IsNullOrWhiteSpace(cAsk))
                {
                    cAsk = "コア価値は？"; // 中立フェイルセーフ
                    Console.WriteLine("[DEBUG] Core iterative question ultimate fallback used (all retries empty)");
                }

                if (!string.IsNullOrWhiteSpace(cAsk))
                {
                    state.History.Add($"Assistant: {cAsk}");
                    TrimHistory(state.History, 30);
                    state.Step4QuestionCount++;

                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }

                    Console.WriteLine($"[USER_MESSAGE] {cAsk}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(cAsk, cAsk), cancellationToken);
                }
                return;
            }

            // Step5（制約事項：文字数 / 文化・配慮 / 法規・レギュレーション / その他）フロー
            if (state.Step1Completed && state.Step2Completed && state.Step3Completed && state.Step4Completed && !state.Step5Completed)
            {
                var consEval = await EvaluateConstraintsAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, state.FinalCoreValue, cancellationToken);
                if (consEval != null) state.LastConstraintsReason = consEval.Reason;
                if (consEval != null && consEval.IsSatisfied)
                {
                    state.ConstraintCharacterLimit = string.IsNullOrWhiteSpace(consEval.CharacterLimit) ? state.ConstraintCharacterLimit : consEval.CharacterLimit;
                    state.ConstraintCultural = string.IsNullOrWhiteSpace(consEval.Cultural) ? state.ConstraintCultural : consEval.Cultural;
                    state.ConstraintLegal = string.IsNullOrWhiteSpace(consEval.Legal) ? state.ConstraintLegal : consEval.Legal;
                    state.ConstraintOther = string.IsNullOrWhiteSpace(consEval.Other) ? state.ConstraintOther : consEval.Other;

                    var summary = RenderConstraintSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {summary}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(summary, summary), cancellationToken);

                    state.Step5Completed = true; // 制約事項確定

                    var consolidatedAfterStep5 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep5}");
                    await SendSummaryAsync(turnContext, consolidatedAfterStep5, cancellationToken);

                    // 直後に Step6 (要約) → Step7 (クリエイティブ要素) を自動実行
                    if (!state.Step6Completed && _kernel != null)
                    {
                        await TryStep6SummaryAsync(turnContext, state, cancellationToken);
                    }
                    if (state.Step6Completed && !state.Step7Completed)
                    {
                        await TryStep7CreativeAsync(turnContext, state, cancellationToken);
                    }
                    if (state.Step6Completed && state.Step7Completed && !state.Step8Completed)
                    {
                        var reset = await TryStep8ExcelAsync(turnContext, state, cancellationToken);
                        if (reset) return; // 新セッション開始済み
                    }

                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 次の制約確認質問
                var nextConsQ = await GenerateNextConstraintsQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, state.FinalCoreValue, state.Step5QuestionCount, state.LastConstraintsReason, cancellationToken);
                if (string.IsNullOrWhiteSpace(nextConsQ))
                {
                    nextConsQ = await GenerateNextConstraintsQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, state.FinalCoreValue, state.Step5QuestionCount, state.LastConstraintsReason, cancellationToken);
                }
                if (!string.IsNullOrWhiteSpace(nextConsQ))
                {
                    state.History.Add($"Assistant: {nextConsQ}");
                    TrimHistory(state.History, 30);
                    state.Step5QuestionCount++;
                    if (_userState != null)
                    {
                        var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    Console.WriteLine($"[USER_MESSAGE] {nextConsQ}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(nextConsQ, nextConsQ), cancellationToken);
                }
                return;
            }

            // Step7（クリエイティブ要素）自動実行フォロー（制約確定後に要約だけ済んでいて、要素未生成の場合）
            if (state.Step1Completed && state.Step2Completed && state.Step3Completed && state.Step4Completed && state.Step5Completed && state.Step6Completed && !state.Step7Completed)
            {
                await TryStep7CreativeAsync(turnContext, state, cancellationToken);
                if (state.Step7Completed && !state.Step8Completed)
                {
                    var reset2 = await TryStep8ExcelAsync(turnContext, state, cancellationToken);
                    if (reset2) return;
                }
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }

            // Step1: 目的評価
            var eval = await EvaluatePurposeAsync(_kernel, transcript, cancellationToken);
            if (eval != null) state.LastPurposeReason = eval.Reason;
            if (eval != null && eval.IsSatisfied)
            {
                // 目的が十分に引き出せたと判断された場合

                state.FinalPurpose = string.IsNullOrWhiteSpace(eval.Purpose) ? state.FinalPurpose : eval.Purpose;
                var purposeText = state.FinalPurpose ?? eval.Purpose ?? "(未取得)";

                // 要約生成（JSON）→ ユーザー向け整形
                var summaryJson = await GeneratePurposeSummaryAsync(_kernel, transcript, purposeText, cancellationToken);
                state.Step1SummaryJson = summaryJson;
                var summaryView = RenderPurposeSummaryForUser(summaryJson, fallbackPurpose: purposeText);

                // 別バブルで送信: 1) 承知メッセージ（AI生成） 2) 要約
                var ackMsg = await GenerateAckMessageAsync(_kernel, transcript, purposeText, cancellationToken);
                if (string.IsNullOrWhiteSpace(ackMsg))
                {
                    ackMsg = "承知しました。ありがとうございます。"; // フォールバックのみ最小限
                }
                Console.WriteLine($"[USER_MESSAGE] {ackMsg}");
                await turnContext.SendActivityAsync(MessageFactory.Text(ackMsg, ackMsg), cancellationToken);

                var summaryMsg = BuildConsolidatedSummary(state);
                Console.WriteLine($"[USER_MESSAGE] {summaryMsg}");
                await SendSummaryAsync(turnContext, summaryMsg, cancellationToken);

                // Step1完了フラグ
                state.Step1Completed = true;

                // 直後にStep2（ターゲット）への最初の短い質問を1つだけ行う
                var tFirst = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, state.LastTargetReason, cancellationToken);
                if (string.IsNullOrWhiteSpace(tFirst))
                {
                    tFirst = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, state.LastTargetReason, cancellationToken);
                }
                if (!string.IsNullOrWhiteSpace(tFirst))
                {
                    state.History.Add($"Assistant: {tFirst}");
                    TrimHistory(state.History, 30);
                    state.Step2QuestionCount++;

                    Console.WriteLine($"[USER_MESSAGE] {tFirst}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(tFirst, tFirst), cancellationToken);
                }

                // 状態保存
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }

            // 未確定: Elicitor に次の質問を生成（Step1用）
            var ask = await GenerateNextQuestionAsync(_kernel, transcript, state.Step1QuestionCount, state.LastPurposeReason, cancellationToken);

            if (string.IsNullOrWhiteSpace(ask))
            {
                // 1回だけAIに再試行（テンプレ固定文なし）
                ask = await GenerateNextQuestionAsync(_kernel, transcript, state.Step1QuestionCount, state.LastPurposeReason, cancellationToken);
            }

            if (!string.IsNullOrWhiteSpace(ask))
            {
                state.History.Add($"Assistant: {ask}");
                TrimHistory(state.History, 30);
                // 質問回数をカウントして過度な深掘りを避ける
                state.Step1QuestionCount++;

                // 状態保存
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }

                Console.WriteLine($"[USER_MESSAGE] {ask}");
                await turnContext.SendActivityAsync(MessageFactory.Text(ask, ask), cancellationToken);
            }
            return;
        }
        catch (Exception ex)
        {
            var errorMessage = $"エラーが発生しました: {ex.Message}";
            Console.WriteLine($"[USER_MESSAGE] {errorMessage}");
            await turnContext.SendActivityAsync(MessageFactory.Text(errorMessage, errorMessage), cancellationToken);
        }
    }

    protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
    {
        //初回接続時のメッセージ
        // conversationUpdate の中でも、ユーザーが追加されたときのみウェルカムを一度だけ送る
        var activity = turnContext.Activity;
        var membersAdded = activity.MembersAdded;
        var botId = activity.Recipient?.Id;

        bool userJoined = membersAdded != null && membersAdded.Any(m => m.Id != null && m.Id != botId);
        if (userJoined && !IsWelcomeDisabled())
        {
            var welcomeMessage = UnifiedWelcomeMessage;

            if (_userState != null)
            {
                var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                var state = await accessor.GetAsync(turnContext, () => ElicitationState.CreateNew(), cancellationToken);
                if (!state.WelcomeSent)
                {
                    // 直前にフォールバックで既に送っているケースを保険的に検知（履歴末尾と比較）
                    var lastAssistant = state.History.LastOrDefault(h => h.StartsWith("Assistant:"));
                    string Normalize(string s) => Regex.Replace(s, "\\s+", " ").Trim();
                    var normalizedExisting = lastAssistant is null ? string.Empty : Normalize(lastAssistant);
                    var normalizedWelcome = Normalize("Assistant: " + welcomeMessage);
                    if (!string.Equals(normalizedExisting, normalizedWelcome, StringComparison.Ordinal))
                    {
                        Console.WriteLine($"[USER_MESSAGE] {welcomeMessage}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(welcomeMessage, welcomeMessage), cancellationToken);
                        state.History.Add($"Assistant: {welcomeMessage}");
                        TrimHistory(state.History, 30);
                    }
                    state.WelcomeSent = true;
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
            }
            else
            {
                // ConversationState 未設定の場合はガードできないため、membersAdded条件のみで送る
                Console.WriteLine($"[USER_MESSAGE] {welcomeMessage}");
                await turnContext.SendActivityAsync(MessageFactory.Text(welcomeMessage, welcomeMessage), cancellationToken);
            }
        }

        await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);
    }

    private static void TrimHistory(List<string> history, int max)
    {
        if (history.Count > max)
        {
            history.RemoveRange(0, history.Count - max);
        }
    }

    // ユーザーが「次へ進みたい/スキップしたい」明示を含むか簡易判定
    private static bool UserWantsAdvance(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var t = text.Trim();
        string[] patterns = { "次へ", "次に進", "次いこ", "次行", "次のステップ", "先へ", "先に進", "スキップ", "skip", "もういい", "大丈夫です", "進んで", "進めて", "省略", "飛ばして" };
        foreach (var p in patterns)
        {
            if (t.Contains(p, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    // 目的受領時の承知メッセージをAIに1文で生成させる
    private static async Task<string?> GenerateAckMessageAsync(Kernel kernel, string transcript, string purpose, CancellationToken ct)
    {
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは丁寧で軽やかなアシスタントです。以下の会話と受け取った活動目的を踏まえ、
承知の意を短く1文だけ日本語で伝えてください。敬語は自然体で、堅すぎず、命令形や謝罪は避けます。
禁止：目的の言い換えの羅列、評価的コメント、次の質問。
出力はそのままユーザーに見せる1文のみ。");
        history.AddUserMessage($"--- 受領目的 ---\n{purpose}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S1/PURPOSE:ACK");
        return response?.Content?.Trim();
    }

    // Step2: ターゲットが十分に定義されたかを評価
    private static async Task<TargetDecision?> EvaluateTargetAsync(Kernel kernel, string transcript, string? step1SummaryJson, CancellationToken ct)
    {
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ2『ターゲット』のみを評価します。
            ここでいうターゲットは、誰に届けるかの像（例：属性、役割、状況、顧客段階など）です。
            会話履歴と、もしあればステップ1要約JSON内の 'audience' や 'references.targetHints' を手掛かりに、既出の情報のみから判断します。

            判定基準：
            - 誰に向けたコピーかが1行で説明できること（例：関東圏の大学1〜2年生、既存のライトユーザー など）
            - 既出の事実に基づくこと（推測や新規追加はしない）

            【ユーザー質問優先ルール（再定義）】
            - 対象は『ユーザーが Assistant に情報/説明を求める明示的質問』のみ。
            - 直近発話にその種の未回答質問が含まれ、まだ Assistant が回答していない場合のみ isSatisfied=false。
            - ユーザーがこちらの質問に答えず保留/拒否/スキップ（例: 「まだ」「特にない」「次へ」等）しただけなら未回答扱いにしない。
            - 『次へ』『スキップ』『もういい』など前進希望があれば本ルール適用せず他基準のみで判定。
            - 未回答がある場合 reason を『未回答ユーザー質問: 要点…』で始め1行。無ければ不足指摘のみ。
            - JSON 内でその質問へ直接回答しない。

            出力は次のJSONのみ（先頭や末尾に ``` や ```json などのコードフェンスや説明文を付けない）：
            {
                ""isSatisfied"": true,
                ""target"": ""1行の要約（未確定なら空）"",
                ""reason"": ""判断理由（不足点も簡潔に）""
            }");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

    var response = await InvokeAndLogAsync(kernel, history, ct, "S2/TARGET:EVAL");
        var raw = response?.Content?.Trim();
        var json = ExtractFirstJsonObject(raw);
        if (string.IsNullOrWhiteSpace(json)) return null;
        try {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True;
            string? target = root.TryGetProperty("target", out var tEl) && tEl.ValueKind == JsonValueKind.String ? tEl.GetString() : null;
            string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null;
            return new TargetDecision { IsSatisfied = isSat, Target = target, Reason = reason }; }
        catch { return null; }
    }

    // Step2: 次のターゲット確認質問を生成（既出の手がかりを活かす）
    private static async Task<string?> GenerateNextTargetQuestionAsync(Kernel kernel, string transcript, string? step1SummaryJson, int questionCount, string? evalReason, CancellationToken ct)
    {
        var history = new ChatHistory();
        var pacing = questionCount >= 4
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の選択肢を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        var reasonSnippet = string.IsNullOrWhiteSpace(evalReason) ? string.Empty : $"直前の評価で不足とされた点: {evalReason}\n";
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『ターゲット』だけを確かめます。
            {reasonSnippet}
            既出の手がかり（Step1要約のaudienceやtargetHints、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            ターゲット以外（目的の再評価、表現案、制作条件など）は扱いません（別フェーズ）。

            不足Reasonがある場合はその原文を羅列せず、自然な言い換えで 1 点だけギャップを埋める質問にしてください。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            【未回答質問ブリッジ】(evalReason に「未回答」が含まれる場合のみ)
            - 先に未回答ユーザー質問へ 1 行で簡潔に回答（要点を自然に言い換え、引用丸写し禁止）
            - その流れでターゲット像の不足 1 点に絞った質問を 1 つだけ提示
            - 回答と質問は '。' でつなぐか 1 文にまとめる。複数質問は禁止
            - 回答のみで十分なら「他にターゲット像で補足があれば教えてください。」等で締めてもよい

            出力はユーザーにそのまま見せる日本語のテキストのみ。評価理由のコピーペーストは禁止。");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON（参考） ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

    var response = await InvokeAndLogAsync(kernel, history, ct, "S2/TARGET:Q");
        var tq = response?.Content?.Trim();
        if (!string.IsNullOrWhiteSpace(tq) && questionCount >= 2 && tq.Length < 140 && !tq.Contains("次へ"))
        {
            tq += " （十分であれば『次へ』とだけ返信で先に進みます）";
        }
        return tq;
    }

    // Step3: 媒体/利用シーンが十分に定義されたかを評価
    private static async Task<MediaDecision?> EvaluateMediaAsync(Kernel kernel, string transcript, string? step1SummaryJson, CancellationToken ct)
    {
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ3『媒体／利用シーン』のみを評価します。
            ここでいう媒体／利用シーンは、キャッチコピーがどこで・どのように使われるか（例：LPのヒーロー、駅ポスター、アプリ内バナー、メール件名 等）です。
            会話履歴と、もしあればステップ1要約JSON内の 'usageContext' を手掛かりに、既出の情報のみから判断します。

            判定基準：
            - 媒体／利用シーンが1行で説明できること（例：特設LPのファーストビュー、店頭A1ポスター など）
            - 既出の事実に基づくこと（推測や新規追加はしない）

            【ユーザー質問優先ルール（再定義）】
            - Assistant への説明/判断を求める未回答質問が直近発話に残っている場合のみ isSatisfied=false。
            - 「特にない」「決めてない」「次へ進んで」等は質問ではないため無視。
            - 前進希望（次へ/スキップ等）があれば本ルールを適用せず他基準のみで判定。
            - 未回答がある場合 reason を『未回答ユーザー質問: 要点…』で始め 1 行。無ければ不足/確定理由のみ。

            出力は次のJSONのみ（コードフェンスや説明文禁止）：
            {
                ""isSatisfied"": true,
                ""mediaOrContext"": ""1行の要約（未確定なら空）"",
                ""reason"": ""判断理由（不足点も簡潔に）""
            }");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

    var response = await InvokeAndLogAsync(kernel, history, ct, "S3/MEDIA:EVAL");
        var raw = response?.Content?.Trim();
        var json = ExtractFirstJsonObject(raw);
        if (string.IsNullOrWhiteSpace(json)) return null;
        try { using var doc = JsonDocument.Parse(json); var root = doc.RootElement; bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True; string? media = root.TryGetProperty("mediaOrContext", out var mEl) && mEl.ValueKind == JsonValueKind.String ? mEl.GetString() : null; string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null; return new MediaDecision { IsSatisfied = isSat, MediaOrContext = media, Reason = reason }; } catch { return null; }
    }

    // Step3: 次の媒体確認質問を生成（既出の手がかりを活かす）
    private static async Task<string?> GenerateNextMediaQuestionAsync(Kernel kernel, string transcript, string? step1SummaryJson, int questionCount, string? evalReason, CancellationToken ct)
    {
        var history = new ChatHistory();
        var pacing = questionCount >= 4
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の選択肢を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        var reasonSnippet = string.IsNullOrWhiteSpace(evalReason) ? string.Empty : $"直前の評価で不足とされた点: {evalReason}\n";
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『媒体／利用シーン』だけを確かめます。
            {reasonSnippet}
            既出の手がかり（Step1要約のusageContext、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            媒体以外（目的やターゲットの再確認、表現案、制作条件など）は扱いません（別フェーズ）。

            不足Reasonがある場合は原文丸写しを避け、自然な言い換えで 1 点だけギャップを埋める質問にしてください。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            【未回答質問ブリッジ】(evalReason に「未回答」が含まれる場合のみ)
            - 先に未回答ユーザー質問へ 1 行で簡潔に回答（要点の自然な言い換え）
            - 続けて媒体/利用シーンの不足 1 点を埋める質問を 1 つだけ提示
            - 回答と質問は '。' でつなぐか 1 文にまとめる。複数質問は禁止
            - 回答のみで十分なら「他に媒体や利用シーンで足りない点があれば教えてください。」等で締めてもよい

            出力はユーザーにそのまま見せる日本語のテキストのみ。評価理由のコピーペーストは禁止。");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON（参考） ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S3/MEDIA:Q");
        var mq = response?.Content?.Trim();
        if (!string.IsNullOrWhiteSpace(mq) && questionCount >= 2 && mq.Length < 140 && !mq.Contains("次へ"))
        {
            mq += " （十分であれば『次へ』とだけ返信で先に進みます）";
        }
        return mq;
    }

    // Step6: キャッチコピー制作のためのクリエイティブ要素を自動生成（ユーザー追加入力なし）
    private static async Task<string?> GenerateCreativeElementsAsync(
        Kernel kernel,
        string transcript,
        string? step1SummaryJson,
        string? finalPurpose,
        string? finalTarget,
        string? finalUsageContext,
        string? finalCoreValue,
        string? constraintCharacterLimit,
        string? constraintCultural,
        string? constraintLegal,
        string? constraintOther,
        CancellationToken ct)
    {
        // === 新方式: 直接 JSON を生成させる ===
        var history = new ChatHistory();
        // カテゴリ部分を中央定義から動的生成（変更時にプロンプトとの不整合を避ける）
        var schemaSampleInner = string.Join(",\n            ", CreativeElementCategories.All
            .Select(c => $"\"{c}\": [ \"例1\", \"例2\", \"例3\", \"例4\", \"例5\" ]"));
        var schemaSample = "{" + "\n            " + schemaSampleInner + "\n        }";
        // どのカテゴリが実行時に読み込まれているか明示ログ
    Console.WriteLine("[DEBUG][S7] Categories in runtime: " + string.Join(" | ", CreativeElementCategories.All));
        history.AddSystemMessage($@"""
        あなたは日本語コピーライティング支援アシスタントです。以下の確定情報を踏まえて
        『要素アイデア』を JSON 形式のみで出力してください。余計な説明やコードフェンス、前後テキストは禁止です。

        JSON スキーマ（厳守）例:
        {schemaSample}

        必須カテゴリ（欠落・改名・並び替え禁止。全て 5 要素 埋めること。欠けたら不正）:
        {string.Join(" / ", CreativeElementCategories.All)}

        ガイド:
        - 既出情報と矛盾する具体名や数字を作らない
        - 類似や言い換え重複を避け幅を出す
        - 各配列は必ず 5 要素（不足や過剰禁止）
        - 並び順は意味的に自然なら自由
        - プアな発想は避けること(例：駅前にポスター貼る。では、場所は駅前だ。 ※この場合、駅前にポスターが張られるだけであり、イベントが駅前とはだれも言っていない)

        出力は上記 JSON オブジェクトそのもの。先頭/末尾の空行、説明文、コードフェンス禁止。
        """);

    var ctxLines = new List<string>();
        if (!string.IsNullOrWhiteSpace(finalPurpose)) ctxLines.Add($"目的: {finalPurpose}");
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctxLines.Add($"ターゲット: {finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctxLines.Add($"媒体: {finalUsageContext}");
        if (!string.IsNullOrWhiteSpace(finalCoreValue)) ctxLines.Add($"コア価値: {finalCoreValue}");
    if (!string.IsNullOrWhiteSpace(constraintCharacterLimit)) ctxLines.Add($"文字数制限: {constraintCharacterLimit}");
    if (!string.IsNullOrWhiteSpace(constraintCultural)) ctxLines.Add($"文化/配慮: {constraintCultural}");
    if (!string.IsNullOrWhiteSpace(constraintLegal)) ctxLines.Add($"法規/必須表記: {constraintLegal}");
    if (!string.IsNullOrWhiteSpace(constraintOther)) ctxLines.Add($"その他制約: {constraintOther}");
        if (!string.IsNullOrWhiteSpace(step1SummaryJson)) ctxLines.Add($"(Step1要約JSONあり)");

        history.AddUserMessage($"--- 確定情報 ---\n{string.Join("\n", ctxLines)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S7/ELEMENTS:GEN");
        var raw = response?.Content?.Trim();
        if (string.IsNullOrWhiteSpace(raw)) return raw;

        // 1) 直接 JSON パース試行
        Dictionary<string, List<string>>? dict = TryParseCreativeElementsJson(raw);
        if (dict == null)
        {
            // 2) フェンス等除去再試行
            var cleaned = StripFences(raw);
            dict = TryParseCreativeElementsJson(cleaned);
        }
        if (dict == null)
        {
            // 3) 旧スタイル（箇条書き）フォールバックパース
            dict = ParseCreativeElementsToSimpleJson(raw);
        }

        if (dict != null)
        {
            try
            {
                var json = JsonSerializer.Serialize(dict, new JsonSerializerOptions { WriteIndented = true });
                var exportDir = Path.Combine(AppContext.BaseDirectory, "exports");
                Directory.CreateDirectory(exportDir);
                var fileName = $"creative_elements_{DateTime.UtcNow:yyyyMMdd_HHmmss}.json";
                var path = Path.Combine(exportDir, fileName);
                File.WriteAllText(path, json, System.Text.Encoding.UTF8);
                Console.WriteLine($"[STEP7_JSON_SAVED] {path}");
                // 人間向けにレンダリングして返す
                return RenderCreativeElementsForUser(dict);
            }
            catch (Exception exWrite)
            {
                Console.WriteLine($"[STEP7_JSON_WRITE_ERROR] {exWrite.Message}");
                return RenderCreativeElementsForUser(dict); // 保存失敗でも表示は行う
            }
        }
    Console.WriteLine("[STEP7_JSON_TOTAL_FAIL] JSON/旧形式いずれも解析不可 -> 生raw返却");
        return raw; // 最悪そのまま
    }


    // Step4: 提供価値・差別化のコアが十分に定義されたかを評価
    private static async Task<CoreDecision?> EvaluateCoreAsync(
        Kernel kernel,
        string transcript,
        string? step1SummaryJson,
        string? finalTarget,
        string? finalUsageContext,
        CancellationToken ct)
    {
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ4『提供価値・差別化のコア』のみを評価します。
            ここでいうコアとは、ユーザーにとっての価値の中核や、競合と比べたときの差別化要素です。
            会話履歴と、もしあればStep1要約JSON内の 'coreDriver' や 'subjectOrHero'、'references.essenceHints' を手掛かりに、既出の情報のみから判断します。

            判定基準：
            - どんな価値／差別化を打ち出すのかが1行で説明できること（例：初心者でも10分で設定完了、地元の実例ストーリーで信頼感、等）
            - 既出の事実に基づくこと（推測や新規追加はしない）

            【ユーザー質問優先ルール（再定義）】
            - 直近ユーザー発話に Assistant への回答要求を含む未回答質問が残っている場合のみ isSatisfied=false。
            - 回答拒否/保留/スキップや『次へ』は未回答質問とは見なさない。
            - 未回答がある場合 reason を『未回答ユーザー質問: 要点…』で始め 1 行。無ければ不足理由のみ。

            出力は次のJSONのみ（コードフェンス禁止）：
            {
                ""isSatisfied"": true,
                ""core"": ""1行の要約（未確定なら空）"",
                ""reason"": ""判断理由（不足点も簡潔に）""
            }");

        var ctx = new List<string>();
        ctx.Add(string.IsNullOrWhiteSpace(step1SummaryJson) ? "(Step1要約なし)" : step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[FinalTarget] {finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[FinalUsageContext] {finalUsageContext}");

        history.AddUserMessage($"--- コンテキスト ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S4/CORE:EVAL");
        var raw = response?.Content?.Trim();
        var json = ExtractFirstJsonObject(raw);
        if (string.IsNullOrWhiteSpace(json)) return null;
        try { using var doc = JsonDocument.Parse(json); var root = doc.RootElement; bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True; string? core = root.TryGetProperty("core", out var cEl) && cEl.ValueKind == JsonValueKind.String ? cEl.GetString() : null; string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null; return new CoreDecision { IsSatisfied = isSat, Core = core, Reason = reason }; } catch { return null; }
    }

    // === Step5（制約事項）関連 ヘルパー ===
    // 以前は初期質問ハードコード / パターン判定が存在したが、ユーザー要望で撤去。
    // 初期質問も含め LLM プロンプトの記述力に委ねる設計へ移行。

    private class ConstraintDecision
    {
        public bool IsSatisfied { get; set; }
        public string? CharacterLimit { get; set; }
        public string? Cultural { get; set; }
        public string? Legal { get; set; }
        public string? Other { get; set; }
        public string? Reason { get; set; }
    }

    private static async Task<ConstraintDecision?> EvaluateConstraintsAsync(
        Kernel kernel,
        string transcript,
        string? step1SummaryJson,
        string? finalTarget,
        string? finalUsageContext,
        string? finalCoreValue,
        CancellationToken ct)
    {
        var history = new ChatHistory();
                history.AddSystemMessage(@"あなたは第三者レビュアーです。以下の会話から '制約事項'（キャッチコピー制作時に守るべき条件）が十分か判定します。
                制約カテゴリ: 1) 文字数制限 2) 文化・配慮事項 3) 法規・レギュレーション / 必須表記 4) その他（NGワードや内部基準など）
                既出情報のみを使用し、推測で新規追加しない。曖昧・未言及は空文字。
                判定基準: 4カテゴリすべてにおいて『明確に不要』または『明確に記述あり』の状態なら isSatisfied=true。
                
                【ユーザー質問優先ルール（再定義）】
                - 直近ユーザー発話に Assistant への説明/判断を求める未回答質問が残っている場合のみ isSatisfied=false。
                - 『特にない / まだ / わからない / 次へ / スキップ』等は質問ではないため適用しない。
                - 前進希望（次へ/スキップ等）があれば本ルールを無視して他基準で判定。
                - 未回答がある場合 reason を『未回答ユーザー質問: 要点…』で始め 1 行。無い場合は通常理由。
                出力JSONのみ（コードフェンス禁止）:
                {
                    ""isSatisfied"": true,
                    ""characterLimit"": ""例: 全角15文字以内 / 指定なし"",
                    ""cultural"": ""例: 差別的ニュアンス回避 / 指定なし"",
                    ""legal"": ""例: 医療効能表現禁止 / 指定なし"",
                    ""other"": ""例: 社内用語は避ける / 指定なし"",
                    ""reason"": ""簡潔な判定根拠""
                }");
        var ctx = new List<string>();
        if (!string.IsNullOrWhiteSpace(step1SummaryJson)) ctx.Add(step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[Target]{finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[Usage]{finalUsageContext}");
        if (!string.IsNullOrWhiteSpace(finalCoreValue)) ctx.Add($"[Core]{finalCoreValue}");
        history.AddUserMessage($"--- コンテキスト ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S5/CONSTRAINTS:EVAL");
        var raw = response?.Content?.Trim();
        var json = ExtractFirstJsonObject(raw);
        if (string.IsNullOrWhiteSpace(json)) return null;
        try { using var doc = JsonDocument.Parse(json); var root = doc.RootElement; bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True; string GetStr(string name) => root.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.String ? (el.GetString() ?? "") : ""; return new ConstraintDecision { IsSatisfied = isSat, CharacterLimit = GetStr("characterLimit"), Cultural = GetStr("cultural"), Legal = GetStr("legal"), Other = GetStr("other"), Reason = GetStr("reason") }; } catch { return null; }
    }

    private static async Task<string?> GenerateNextConstraintsQuestionAsync(
        Kernel kernel,
        string transcript,
        string? step1SummaryJson,
        string? finalTarget,
        string? finalUsageContext,
        string? finalCoreValue,
        int questionCount,
        string? evalReason,
        CancellationToken ct)
    {
        var history = new ChatHistory();
    var pacing = questionCount >= 4 ? "（質問が続いているため簡潔に。YES/NOか選択肢提示で手短に）" : string.Empty;
        var reasonSnippet = string.IsNullOrWhiteSpace(evalReason) ? string.Empty : $"直前の評価で不足とされた点: {evalReason}\n";
        history.AddSystemMessage(@$"あなたは丁寧なコピー制作アシスタントです。今は『制約事項』（文字数 / 文化・配慮 / 法規・レギュレーション / その他）だけを確認します。
        {reasonSnippet}
        既出を繰り返しすぎない。新規に想像で条件を作らない。{pacing}

        厳守ルール:
        - 出力は必ず 1 つの質問文（疑問符 ? または ？ を含む）
        - 'YES' や 'はい' などの了承単語のみを返してはいけない
        - 箇条書きは最大 1 行に留め、冗長な前置き禁止
        - まだ未確定/空のカテゴリだけを明示的に軽く聞くのは OK
        - 不足Reasonが示すカテゴリのみ絞り込んで確認（複数同時に羅列しない）
        - 追加が無ければ『特になし』と答えてください、のような誘導を含めてもよい

        出力は質問文 1 行のみ（説明・コードフェンス禁止）。評価理由のコピーペーストは禁止。");
        history.AddSystemMessage(@"【未回答質問ブリッジ】(evalReason に『未回答』が含まれる場合のみ)
            - 冒頭で未回答ユーザー質問へ 1 行で簡潔に回答（要点言い換え）
            - 同一行で不足する制約 1 点のみを尋ねる質問（? / ？ を含む）を続ける
            - 全体を 1 行に保ち複数質問禁止。冗長説明禁止");
        var ctx = new List<string>();
        if (!string.IsNullOrWhiteSpace(step1SummaryJson)) ctx.Add(step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[Target]{finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[Usage]{finalUsageContext}");
        if (!string.IsNullOrWhiteSpace(finalCoreValue)) ctx.Add($"[Core]{finalCoreValue}");
        history.AddUserMessage($"--- コンテキスト（参考） ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S5/CONSTRAINTS:Q");
        var cq = response?.Content?.Trim();
        if (!string.IsNullOrWhiteSpace(cq) && questionCount >= 2 && cq.Length < 140 && !cq.Contains("次へ"))
        {
            cq += " （十分であれば『次へ』とだけ返信で先に進みます）";
        }
        return cq;
    }
    // (削除済) IsInvalidConstraintQuestion / IsNoConstraintsUtterance / 固定フォールバック / 誘導文
    // すべて撤去し、回答スタイルはプロンプト上の指示とモデル判断に任せる。

    private static string RenderConstraintSummary(ElicitationState state)
    {
        var lines = new List<string>();
        lines.Add("— 制約事項まとめ —");
        lines.Add($"文字数: {(!string.IsNullOrWhiteSpace(state.ConstraintCharacterLimit) ? state.ConstraintCharacterLimit : "指定なし")}");
        lines.Add($"文化・配慮: {(!string.IsNullOrWhiteSpace(state.ConstraintCultural) ? state.ConstraintCultural : "指定なし")}");
        lines.Add($"法規・レギュレーション: {(!string.IsNullOrWhiteSpace(state.ConstraintLegal) ? state.ConstraintLegal : "指定なし")}");
        lines.Add($"その他: {(!string.IsNullOrWhiteSpace(state.ConstraintOther) ? state.ConstraintOther : "指定なし")}");
        // クライアントによっては \n だけだと単一行扱いされる可能性があるため CRLF で連結
        var summary = string.Join("\r\n", lines);
        // デバッグ: 実際の可視化用に制御文字をエスケープした形をログ出力
        try
        {
            var visible = summary.Replace("\r", "<CR>").Replace("\n", "<LF>");
            Console.WriteLine($"[SUMMARY_DEBUG] raw='{visible}' len={summary.Length}");
        }
        catch { /* ignore logging errors */ }
        return summary;
    }

    // Step4: 次のコア確認質問を生成（既出の手がかりを活かす）
    private static async Task<string?> GenerateNextCoreQuestionAsync(
        Kernel kernel,
        string transcript,
        string? step1SummaryJson,
        string? finalTarget,
        string? finalUsageContext,
        int questionCount,
        string? evalReason,
        CancellationToken ct)
    {
        var history = new ChatHistory();
        var pacing = questionCount >= 4
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の仮説候補を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        var reasonSnippet = string.IsNullOrWhiteSpace(evalReason) ? "" : $"直前の評価で不足とされた点: {evalReason}\n";
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『提供価値・差別化のコア』だけを確かめます。
            {reasonSnippet}
            既出の手がかり（Step1要約のcoreDriver/subjectOrHero/essenceHints、確定済みのターゲットや媒体、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            コア以外（目的/ターゲット/媒体の再確認、表現案、制作条件など）は扱いません（別フェーズ）。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            出力はユーザーにそのまま見せる日本語のテキストのみ。評価理由をそのまま繰り返し羅列するのは避け、質問に自然に反映してください。");
        history.AddSystemMessage(@"【未回答質問ブリッジ】(evalReason に『未回答』が含まれる場合のみ)
            - 先に未回答ユーザー質問へ 1 行で簡潔に回答（要点言い換え）
            - 続けてコア価値の不足 1 点にフォーカスした質問を 1 つ
            - 回答と質問は '。' でつなぐか 1 文にまとめる。複数質問禁止
            - 回答のみで十分なら『他にコアの価値で補足があれば教えてください。』等で締めてもよい");
        history.AddSystemMessage(@"【出力フォーマット厳格化】
            - 必ず 1 つの日本語の質問文を返す（疑問符 ? または ？ を含む）
            - 空/回答だけ/列挙羅列/二つ以上の質問/挨拶のみ は不可
            - 30文字以内目安で一点に集中
            - 生成不能でも空文字は禁止（最低限の質問を返す）");

        var ctx = new List<string>();
        ctx.Add(string.IsNullOrWhiteSpace(step1SummaryJson) ? "(Step1要約なし)" : step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[FinalTarget] {finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[FinalUsageContext] {finalUsageContext}");
        history.AddUserMessage($"--- コンテキスト（参考） ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S4/CORE:Q");
        var cq = response?.Content?.Trim();
        if (!string.IsNullOrWhiteSpace(cq) && questionCount >= 2 && cq.Length < 140 && !cq.Contains("次へ"))
        {
            cq += " （十分であれば『次へ』とだけ返信で先に進みます）";
        }
        return cq;
    }


        private static async Task<EvalDecision?> EvaluatePurposeAsync(Kernel kernel, string transcript, CancellationToken ct)
    {
    var history = new ChatHistory();
                history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ1『活動目的（なぜ作るのか）』のみを評価します。
                ターゲット・表現案・制作条件などは扱いません。まず『どんな活動のためか』と、それがイベント / 販促キャンペーン / ブランド認知 である場合は『何の（どの）イベント・製品/サービス・ブランドか』が特定できているかに注目します（採用活動のみは職種や人数など未提示でも目的性が明確なら可）。

                主なカテゴリ例：
                - イベント告知（例：地域音楽フェス、学園祭、◯◯展示会 などイベントの種類/名称が分かる）
                - 販促キャンペーン（例：新発売の◯◯飲料、既存アプリのプレミアムプラン、季節限定◯◯商品 など対象が分かる）
                - ブランド認知（例：◯◯という新ブランド、◯◯サービスの信頼向上 など対象領域が分かる）
                - 採用（目的性自体が十分具体：新卒採用強化 / エンジニア中途採用 など。採用は対象詳細が薄くても目的として成立し得る）
                - その他（上記以外）

                追加の厳格化:
                - ｢イベントを告知したい｣ だけで『何のイベントか（種類 or 名称）』が無い場合は不足として isSatisfied=false。
                - ｢販促したい / キャンペーンを打ちたい｣ だけで『何を（製品/サービス/プラン等）』が無いなら isSatisfied=false。
                - ｢ブランド認知を広げたい｣ だけで『どのブランド/サービス領域』が無いなら isSatisfied=false。
                - 採用は単語だけ（例: 採用活動）でも一次的に isSatisfied=true 可。ただし具体職種等があれば purpose に含めてよい。

                判定基準（全て満たすとき isSatisfied=true）:
                1. 活動カテゴリが明確（上記のどれか or 類似）
                2. 上記で追加特定が必須とされたカテゴリでは対象（イベント種類/名称, 製品/サービス, ブランド領域）が一文内で把握できる
                3. 一文で簡潔に目的を言い切れる

                不足があれば isSatisfied=false とし reason に不足点（例: イベント種類不明 / 対象製品不明 など）を列挙。

                【ユーザー質問優先ルール（再定義）】
                - Assistant への情報/説明を求める未回答質問が直近にある場合のみ isSatisfied=false。
                - ユーザーがこちらの質問に答えていない/保留しただけ（『まだ』『特にない』『後で』『次へ』等）は未回答扱いにしない。
                - 前進希望（次へ/スキップ等）があれば本ルールは適用せず他基準で判定。
                - 未回答がある場合 reason を『未回答ユーザー質問: 要点…』で始め 1 行。無い場合は不足指摘のみ。

                最初の一回目の会話であれば、ユーザーは、どうすればいいかわかっていない可能性があるので、何をする場面かを説明するようにしてください。

                出力は次のJSONのみ。コードフェンスや余計な前置きは禁止:
                {
                    ""isSatisfied"": true,
                    ""purpose"": ""活動目的を短く一文で。未確定なら空でもよい"",
                    ""reason"": ""判断理由を簡潔に（不足があれば指摘）""
                }");
        history.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S1/PURPOSE:EVAL");
        var raw = response?.Content?.Trim();
        var json = ExtractFirstJsonObject(raw);
        if (string.IsNullOrWhiteSpace(json)) return null;
        try { using var doc = JsonDocument.Parse(json); var root = doc.RootElement; bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True; string? purpose = root.TryGetProperty("purpose", out var pEl) && pEl.ValueKind == JsonValueKind.String ? pEl.GetString() : null; string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null; return new EvalDecision { IsSatisfied = isSat, Purpose = purpose, Reason = reason }; } catch { return null; }
    }

    private static async Task<string?> GenerateNextQuestionAsync(Kernel kernel, string transcript, int questionCount, string? evalReason, CancellationToken ct)
    {
        var history = new ChatHistory();
        var pacing = questionCount >= 4
            ? "（すでに質問が続いているため、深掘りは控えめに。2〜3個の方向性の仮説を各1行で提案し、『どれが近い？／他にありますか？』とだけ確認。畳みかけはNG）"
            : string.Empty;
        var reasonSnippet = string.IsNullOrWhiteSpace(evalReason) ? string.Empty : $"直前の評価で不足とされた点: {evalReason}\n";
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『活動目的』だけを確かめます。
            テンプレ口調は避け、許可ベースで、短く1つだけ質問してください。質問密度は控えめにしてください{pacing}。

            目的の深度要件（不足があればまずそこを1ステップで聞く）:
            - イベント告知 → 何のイベントか（種類/名称/テーマ）。例: 地域◯◯フェス / 学園祭 / 新製品発表会
            - 販促キャンペーン → 何の製品・サービス・プランか
            - ブランド認知 → どのブランド / どのサービス領域か
            - 採用 → そのままでも可（必要なら職種や層を任意確認）

            既にカテゴリは出ていて “何の◯◯か” が欠けている場合は、それを一問で埋める質問を作成。まだカテゴリ自体が曖昧なら、イベント/販促/ブランド認知/採用/その他 のどれが近いか軽い候補提示も可（最大3行）。

            {reasonSnippet}

            【未回答質問ブリッジ】(evalReason に「未回答」が含まれる場合のみ)
            - まずユーザーの未回答質問へ 1 行で簡潔に回答（質問文の丸写しは禁止。要点を自然に言い換える）
            - 続けて活動目的の不足点 1 つに絞った質問を 1 つだけ提示
            - 回答→質問は '。' で自然につなぐ（必要なら 1 文にまとめる）。複数質問は禁止
            - 回答だけで十分なら質問せず「他に目的で補足があれば教えてください。」等で締めてもよい

            最初の会話であれば、ユーザーは、どうすればいいかわかっていない可能性があるので、何をする場面かを説明するようにしてください。
            ターゲット像や表現案、制作条件には踏み込みません（別フェーズ）。

            不足Reasonがある場合は原文を羅列せず自然な言い換えで 1 点だけ埋める質問にしてください。

            出力はユーザーにそのまま見せる日本語テキストのみ。余計な前置きや工程名は禁止。");
        history.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
        var response = await InvokeAndLogAsync(kernel, history, ct, "S1/PURPOSE:Q");
        var q = response?.Content?.Trim();
        if (!string.IsNullOrWhiteSpace(q) && questionCount >= 2 && q.Length < 140 && !q.Contains("次へ"))
        {
            q += " （十分であれば『次へ』とだけ返信で先に進みます）";
        }
        return q;
    }

    // Step1（目的）サマリーの生成（JSON）
        private static async Task<string> GeneratePurposeSummaryAsync(Kernel kernel, string transcript, string purpose, CancellationToken ct)
    {
                var history = new ChatHistory();
                history.AddSystemMessage(@"あなたは第三者のレビュアー兼まとめ役です。今はステップ1『活動目的』を中心に、会話中に『既に出ている情報』だけを静かに拾って要約（JSON）に整えます。
                    重要: ユーザーに追加質問はしません。推測や創作は厳禁。会話に現れていない項目は空/空配列にしてください。

                    特に、会話中に出ていれば次を拾います：
                    - どんな媒体/利用シーン（例: ランディングページ、ポスター 等）
                    - 対象・主役（例: 製品名、サービス、イベント名 など）
                    - 大枠ターゲット（例: 学生、既存顧客、来場見込み者 など）
                    - 達成したい効果の種類（ゴールイメージ）

                    出力は次のJSONのみ（コードフェンス禁止）：
                    {
                        ""purpose"": ""活動目的を一文で（例：学園祭の来場を増やすためのイベント告知）"",
                        ""usageContext"": ""（任意）媒体/利用シーンが分かれば1行、なければ空"",
                        ""subjectOrHero"": ""（任意）対象・主役が分かれば1行、なければ空"",
                        ""audience"": ""（任意）大枠ターゲットが分かれば1行、なければ空"",
                        ""finalRole"": ""（空で可）"",
                        ""expectedAction"": ""（任意）達成したい効果/期待する行動（例: 訪問/申込/来場 等）"",
                        ""oneSentenceGoal"": ""（任意）ゴールイメージを短い一文に"",
                        ""coreDriver"": ""（空で可）"",
                        ""mustInclude"": [],
                        ""mustAvoid"": [],
                        ""timingOrConstraints"": """",
                        ""references"": { ""targetHints"": [], ""essenceHints"": [], ""expressionHints"": [] },
                        ""gaps"": [],
                        ""confidence"": 0.0,
                        ""reviewer"": { ""evaluation"": ""短い所見"", ""missingPoints"": [], ""discomfortSignals"": [], ""guidanceForNextAI"": """" }
                    }");
        history.AddUserMessage($"--- 目的（判定済み） ---\n{purpose}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S1/PURPOSE:SUMMARY");
        var raw = response?.Content?.Trim();
        var json = ExtractFirstJsonObject(raw) ?? raw;
        return string.IsNullOrWhiteSpace(json) ? "{}" : json;
    }

    // ユーザー向けレンダリング（JSON→テキスト）
    private static string RenderPurposeSummaryForUser(string? json, string fallbackPurpose)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(json)) return $"目的: {fallbackPurpose}";
            using var doc = JsonDocument.Parse(json);
            var r = doc.RootElement;
            string GetS(string name) => r.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.String ? (el.GetString() ?? "") : "";
            string purpose = GetS("purpose");
            string usage = GetS("usageContext");
            string subject = GetS("subjectOrHero");
            string audience = GetS("audience");
            string role = GetS("finalRole");
            string action = GetS("expectedAction");
            string goal = GetS("oneSentenceGoal");
            string core = GetS("coreDriver");
            string timing = GetS("timingOrConstraints");

            string JoinArr(string obj, string name)
            {
                if (!r.TryGetProperty("references", out var refs) || refs.ValueKind != JsonValueKind.Object) return "";
                if (!refs.TryGetProperty(name, out var arr) || arr.ValueKind != JsonValueKind.Array) return "";
                var items = arr.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.String).Select(e => e.GetString()).Where(s => !string.IsNullOrWhiteSpace(s))!;
                return string.Join("\n- ", items!);
            }

            var targetHints = JoinArr("references", "targetHints");
            var essenceHints = JoinArr("references", "essenceHints");
            var expressionHints = JoinArr("references", "expressionHints");

            List<string> GetArr(string name)
            {
                if (!r.TryGetProperty(name, out var arr) || arr.ValueKind != JsonValueKind.Array) return new List<string>();
                return arr.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.String)
                    .Select(e => e.GetString() ?? "")
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
            }
            var mustInclude = GetArr("mustInclude");
            var mustAvoid = GetArr("mustAvoid");

            var lines = new List<string>();
            lines.Add($"目的: {(string.IsNullOrWhiteSpace(purpose) ? fallbackPurpose : purpose)}");
            if (!string.IsNullOrWhiteSpace(usage)) lines.Add($"媒体／利用シーン: {usage}");
            if (!string.IsNullOrWhiteSpace(subject)) lines.Add($"対象・主役: {subject}");
            if (!string.IsNullOrWhiteSpace(audience)) lines.Add($"大枠ターゲット: {audience}");
            if (!string.IsNullOrWhiteSpace(role)) lines.Add($"最終的な役割: {role}");
            if (!string.IsNullOrWhiteSpace(action)) lines.Add($"期待する行動: {action}");
            if (!string.IsNullOrWhiteSpace(goal)) lines.Add($"到達ゴール（一文）: {goal}");
            if (!string.IsNullOrWhiteSpace(core)) lines.Add($"背景・コアの理由: {core}");
            if (mustInclude.Count > 0) lines.Add("入れたいこと:\n- " + string.Join("\n- ", mustInclude));
            if (mustAvoid.Count > 0) lines.Add("避けたいこと:\n- " + string.Join("\n- ", mustAvoid));
            if (!string.IsNullOrWhiteSpace(timing)) lines.Add($"タイミング・制約: {timing}");
            if (!string.IsNullOrWhiteSpace(targetHints) || !string.IsNullOrWhiteSpace(essenceHints) || !string.IsNullOrWhiteSpace(expressionHints))
            {
                lines.Add("参考情報:");
                if (!string.IsNullOrWhiteSpace(targetHints)) lines.Add($"- ターゲットの手がかり\n- {targetHints}");
                if (!string.IsNullOrWhiteSpace(essenceHints)) lines.Add($"- 本質の手がかり\n- {essenceHints}");
                if (!string.IsNullOrWhiteSpace(expressionHints)) lines.Add($"- 表現案の手がかり\n- {expressionHints}");
            }
            return string.Join("\n", lines);
               }
        catch
        {
            return $"目的: {fallbackPurpose}";
        }
    }

    // 共通: Azure OpenAI(Semantic Kernel Chat) 呼び出しの入出力を色付きでログ出力
    private static async Task<ChatMessageContent?> InvokeAndLogAsync(Kernel kernel, ChatHistory history, CancellationToken ct, string stage)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        // 入力メッセージをステージ+役割で逐次表示
        foreach (var m in history)
        {
            var roleStr = m.Role.ToString();
            var roleTag = roleStr.Equals("user", StringComparison.OrdinalIgnoreCase) ? "User" : (roleStr.Equals("User", StringComparison.OrdinalIgnoreCase) ? "User" : "AI");
            WriteColored($"[{stage}][{roleTag}] {roleStr}: {m.Content}", ConsoleColor.Blue);
        }
        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        if (response != null && !string.IsNullOrWhiteSpace(response.Content))
        {
            WriteColored($"[{stage}][AI] {response.Content}", ConsoleColor.Red);
        }
        return response;
    }

    // 累積サマリー生成（各ステップの確定情報を統合）
    private static string BuildConsolidatedSummary(ElicitationState state)
   {
        string purpose = state.FinalPurpose ?? "";
        string usage = state.FinalUsageContext ?? "";
        string target = state.FinalTarget ?? "";
        string core = state.FinalCoreValue ?? "";
        string charLimit = state.ConstraintCharacterLimit ?? "";
        string cultural = state.ConstraintCultural ?? "";
        string legal = state.ConstraintLegal ?? "";
        string other = state.ConstraintOther ?? "";

        // Step1 JSON から補完 (未確定箇所のみ)
        if (!string.IsNullOrWhiteSpace(state.Step1SummaryJson))
        {
            try
            {
                using var doc = JsonDocument.Parse(state.Step1SummaryJson);
                var r = doc.RootElement;
                string GetS(string name) => r.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.String ? (el.GetString() ?? "") : "";
                if (string.IsNullOrWhiteSpace(purpose)) purpose = GetS("purpose");
                if (string.IsNullOrWhiteSpace(usage)) usage = GetS("usageContext");
                if (string.IsNullOrWhiteSpace(target)) target = GetS("audience");
            }
            catch { /* ignore */ }
        }

        var lines = new List<string>();
    // Markdown クライアントで確実に段落分離されるよう、ヘッダー行の末尾に半角スペース2つ + 空行を挿入
    lines.Add("— ここまでの整理 —  ");
    lines.Add(string.Empty);
        // ラベルを【】で囲む形式に変更
        if (!string.IsNullOrWhiteSpace(purpose)) lines.Add($"【目的】: {purpose}");
        if (state.Step2Completed && !string.IsNullOrWhiteSpace(target)) lines.Add($"【ターゲット】: {target}");
        if (state.Step3Completed && !string.IsNullOrWhiteSpace(usage)) lines.Add($"【媒体/利用シーン】: {usage}");
        if (state.Step4Completed && !string.IsNullOrWhiteSpace(core)) lines.Add($"【コア価値】: {core}");
        if (state.Step5Completed)
        {
            lines.Add("【制約事項】:");
            lines.Add($"- 文字数: {(string.IsNullOrWhiteSpace(charLimit) ? "指定なし" : charLimit)}");
            lines.Add($"- 文化・配慮: {(string.IsNullOrWhiteSpace(cultural) ? "指定なし" : cultural)}");
            lines.Add($"- 法規・レギュレーション: {(string.IsNullOrWhiteSpace(legal) ? "指定なし" : legal)}");
            lines.Add($"- その他: {(string.IsNullOrWhiteSpace(other) ? "指定なし" : other)}");
        }
        return string.Join("\n", lines);
    }

    // サマリーをMarkdownとして送信（チャネル側で改行が潰れないよう TextFormat=markdown を指定）
    private static async Task SendSummaryAsync(ITurnContext turnContext, string summary, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(summary)) return;
        // 改行は CRLF に統一し markdown として送信
        var normalized = summary.Replace("\r\n", "\n").Replace("\n", "\r\n");
        var activity = MessageFactory.Text(normalized, normalized);
        activity.TextFormat = "markdown"; // 対応チャネルで複数行レンダリング
        await turnContext.SendActivityAsync(activity, cancellationToken);
    }

    // Step6: キャッチフレーズ生成向け要約生成（旧Step7）
    private async Task TryStep6SummaryAsync(ITurnContext turnContext, ElicitationState state, CancellationToken cancellationToken)
    {
        // 旧フラグ互換: Step6Completed に反映させる
        if (state.Step6Completed)
        {
            return; // 既に完了
        }
        if (_kernel == null)
        {
            var msg = "Kernel 未初期化のため要約を実行できません。";
            Console.WriteLine($"[USER_MESSAGE] {msg}");
            await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
            return;
        }
        try
        {
            var pre = "Step6: 会話を要約しています (キャッチフレーズ用)...";
            Console.WriteLine($"[S6/SUMMARY][START] {pre}");
            await turnContext.SendActivityAsync(MessageFactory.Text(pre, pre), cancellationToken);

            var transcript = string.Join("\n", state.History);
            var consolidated = BuildConsolidatedSummary(state);
            var prompt = $@"以下はユーザーとの対話ログです。後続でキャッチフレーズ（短く訴求力のある表現）を生成するための要約 JSON を作成してください。
要件:
1. 出力は JSON のみ
2. フィールド: purpose, target, usageContext, coreValue, emotionalTone, constraints, uniqueAngle, keyPhrases (配列 5-10件), brandEssenceCandidates (配列 3件), riskNotes
3. 不明なフィールドは空文字または空配列
4. キャッチフレーズ化を意識し、平易で濃縮されたキーワード中心
--- 対話ログ ---\n{transcript}\n--- 累積整理 ---\n{consolidated}\nJSON:";

            var history = new ChatHistory();
            history.AddUserMessage(prompt);
            var resp = await InvokeAndLogAsync(_kernel, history, cancellationToken, "S6/SUMMARY");
            var content = resp?.Content?.Trim();
            if (string.IsNullOrWhiteSpace(content))
            {
                var warn = "要約生成失敗: 応答が空でした";
                Console.WriteLine($"[S6/SUMMARY][WARN] {warn}");
                await turnContext.SendActivityAsync(MessageFactory.Text(warn, warn), cancellationToken);
                return;
            }
            var first = content.IndexOf('{');
            var last = content.LastIndexOf('}');
            if (first >= 0 && last > first)
            {
                content = content.Substring(first, last - first + 1);
            }
            // 1st pass: JSON として妥当か検証し、不備があれば再整形トライ
            string? normalized = TryNormalizeTaglineSummaryJson(content, out var validationErrors);
            if (normalized == null)
            {
                // 失敗: 元のAI応答をそのまま保存せず、エラー通知
                var errMsg = "要約JSONの解析に失敗しました: " + string.Join("; ", validationErrors);
                Console.WriteLine($"[S6/SUMMARY][ERROR] {errMsg}");
                await turnContext.SendActivityAsync(MessageFactory.Text(errMsg, errMsg), cancellationToken);
                // 再試行ガイド
                var guide = "再試行するには 少し待ってから任意のメッセージを送ってください（自動で再試行します）。";
                await turnContext.SendActivityAsync(MessageFactory.Text(guide, guide), cancellationToken);
                return;
            }

            state.TaglineSummaryJson = normalized;
            state.Step6Completed = true;
            try
            {
                var exportDir = Path.Combine(AppContext.BaseDirectory, "exports");
                Directory.CreateDirectory(exportDir);
                var ts = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
                var path = Path.Combine(exportDir, $"tagline_summary_{ts}.json");
                await File.WriteAllTextAsync(path, normalized, cancellationToken);
                Console.WriteLine($"[STEP6_JSON_SAVED] {path}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"[STEP6_JSON_SAVE_ERROR] {exSave.Message}");
            }
            var done = "Step6完了: 要約JSONを生成しました。";
            Console.WriteLine($"[S6/SUMMARY][DONE] {done}");
            await turnContext.SendActivityAsync(MessageFactory.Text(done, done), cancellationToken);
            // Step7 開始案内
            var step7Intro = "ヒアリングした内容から、キャッチフレーズを生成するための要素を生成します。";
            Console.WriteLine($"[S7/INTRO] {step7Intro}");
            await turnContext.SendActivityAsync(MessageFactory.Text(step7Intro, step7Intro), cancellationToken);
        }
        catch (Exception ex)
        {
            var err = $"要約生成中に例外: {ex.Message}";
            Console.WriteLine($"[S6/SUMMARY][EXCEPTION] {err}\n{ex}" );
            await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
        }
    }

    // Step7 JSON 正規化 & 簡易スキーマ検証
    private static string? TryNormalizeTaglineSummaryJson(string raw, out List<string> errors)
    {
        errors = new List<string>();
        if (string.IsNullOrWhiteSpace(raw)) { errors.Add("empty"); return null; }
        // 余計なバッククォート・コードフェンス除去
        var cleaned = raw.Trim().Trim('`');
        // 改行先頭にある ```json などを排除
        cleaned = Regex.Replace(cleaned, "^```(?:json)?", string.Empty, RegexOptions.IgnoreCase | RegexOptions.Multiline);
        cleaned = cleaned.Replace("```", "");

        JsonDocument? doc = null;
        try
        {
            doc = JsonDocument.Parse(cleaned);
        }
        catch (Exception ex)
        {
            errors.Add("JSON parse error: " + ex.Message);
            return null;
        }
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
        {
            errors.Add("root is not object");
            return null;
        }
        // 必須フィールド定義
        var requiredString = new[] { "purpose", "target", "usageContext", "coreValue", "emotionalTone", "constraints", "uniqueAngle", "riskNotes" };
        var requiredArray = new[] { "keyPhrases", "brandEssenceCandidates" };
        var outObj = new Dictionary<string, object?>();
        foreach (var key in requiredString)
        {
            if (root.TryGetProperty(key, out var el))
            {
                outObj[key] = el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString();
            }
            else
            {
                outObj[key] = ""; // 欠損は空文字
                errors.Add($"missing:{key}");
            }
        }
        foreach (var key in requiredArray)
        {
            if (root.TryGetProperty(key, out var el) && el.ValueKind == JsonValueKind.Array)
            {
                var list = new List<string>();
                foreach (var item in el.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.String)
                    {
                        var v = item.GetString();
                        if (!string.IsNullOrWhiteSpace(v)) list.Add(v!.Trim());
                    }
                }
                outObj[key] = list;
            }
            else
            {
                outObj[key] = new List<string>();
                errors.Add($"missing_or_not_array:{key}");
            }
        }
        // brandEssenceCandidates は 3件以内にトリム
        if (outObj["brandEssenceCandidates"] is List<string> bec && bec.Count > 3)
        {
            outObj["brandEssenceCandidates"] = bec.Take(3).ToList();
        }
        // keyPhrases は 最大10件にトリム
        if (outObj["keyPhrases"] is List<string> kp && kp.Count > 10)
        {
            outObj["keyPhrases"] = kp.Take(10).ToList();
        }
        // errors があっても最低限整形したJSONを返し、呼び出し元で null 判定せず保存する戦略もあるが、ここでは欠損が多い場合は失敗にする閾値を設定
        var criticalMissing = errors.Count(e => e.StartsWith("missing:")) > 5; // 大半欠損は失敗扱い
        if (criticalMissing)
        {
            return null;
        }
        var normalized = JsonSerializer.Serialize(outObj, new JsonSerializerOptions { WriteIndented = true });
        return normalized;
    }

    // Core質問の最低限バリデーション（質問になっているか / 不要な了承だけで終わっていないか）
    private static bool IsInvalidCoreQuestion(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return true;
        var t = text.Trim();
        // 質問記号必須
        if (!t.Contains('？') && !t.Contains('?')) return true;

        // 以前は「ありがとうございます」で始まるだけで弾いていたが、実際には
        // 「ありがとうございます！〜〜を教えてください？」のように有効な質問になるケースが多い。
        // 了承語のみ + 質問なし を弾きたいので、了承語で開始し、かつ疑問符が末尾近くまで存在しないケースのみ排除。
        string[] softStarts = { "了解しました", "承知しました", "わかりました", "ありがとうございます", "ではこれらの魅力を中心に" };
        if (softStarts.Any(ss => t.StartsWith(ss, StringComparison.Ordinal)))
        {
            // 了承語部分を除いた残りに 10 文字以上の質問本文 or 疑問符を含むかをチェック
            var trimmed = softStarts.First(ss => t.StartsWith(ss, StringComparison.Ordinal));
            var rest = t.Substring(trimmed.Length).Trim();
            // rest に疑問符が含まれていれば OK（質問として成立）
            if (!(rest.Contains('？') || rest.Contains('?')))
            {
                // 疑問符無し → 無効
                return true;
            }
        }

        // 「〜していきます。」等 完了宣言で締めて質問意図が無いものを排除
        string[] completionLike = { "考えていきます", "進めます", "作っていきます" };
        if (completionLike.Any(c => t.Contains(c)) && (t.EndsWith("。") || t.EndsWith("！") || t.EndsWith("です。")) && !(t.Contains('？') || t.Contains('?')))
        {
            return true;
        }

        return false; // 上記条件に当てはまらなければ有効
    }

    private static string BuildDynamicCoreFallbackQuestion(string transcript)
    {
        // transcript は 'User:' / 'Assistant:' 行を含む。直近のアシスタント列挙候補を逆走査で拾う。
        var lines = transcript.Split('\n').Reverse();
        var candidateList = new List<string>();
        var enumPattern = new Regex(@"^\s*(?:Assistant:\s*)?(?:[-・]?\s*|\d+[\.、)]\s*)(\S.+)$");
        foreach (var raw in lines)
        {
            if (candidateList.Count >= 6) break; // 最大6件（後で先頭4件使用）
            if (!raw.Contains("Assistant:")) continue; // Assistant発言内のみ対象（過去ユーザー羅列は無視）
            var content = raw.Substring(raw.IndexOf("Assistant:") + 10).Trim();
            // 引用符除去
            content = content.Trim('"', '“', '”', '『', '』');
            var m = enumPattern.Match(content);
            if (!m.Success) continue;
            var core = m.Groups[1].Value.Trim();
            // 列挙風でない平文の長文を弾く（20文字超を除外）
            if (core.Length > 24) continue;
            // 重複除外
            if (candidateList.Contains(core)) continue;
            candidateList.Add(core);
        }
        // 改善: 2件以上候補があれば列挙形式で聞き直し、1件以下なら段階的フォールバック
        if (candidateList.Count >= 2)
        {
            var condensed = candidateList.Take(4).ToList();
            Console.WriteLine($"[DEBUG] Core fallback enumerated ({candidateList.Count} candidates) -> {string.Join(" / ", condensed)}");
            // 直接列挙を再提示せず LLM 側の次回生成に任せるため空を返す（上位呼び出しで再試行フローがあれば利用）
            return string.Empty;
        }

        // 0 or 1 件しか拾えなかった場合も固定文を避け、LLM の通常質問生成ロジックに委ねる -> 空
        Console.WriteLine($"[DEBUG] Core fallback generic suppressed (candidates={candidateList.Count})");
        return string.Empty;
    }


            // 新 Step7: クリエイティブ要素生成（旧 Step6 ロジック呼び出しをラップ）
            private async Task TryStep7CreativeAsync(ITurnContext turnContext, ElicitationState state, CancellationToken cancellationToken)
            {
                if (state.Step7Completed) return;
                if (_kernel == null)
                {
                    var msg = "Kernel 未初期化のためクリエイティブ要素生成を実行できません。";
                    Console.WriteLine($"[USER_MESSAGE] {msg}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
                    return;
                }
                var transcript = string.Join("\n", state.History);
                var gen = await GenerateCreativeElementsAsync(
                    _kernel,
                    transcript,
                    state.Step1SummaryJson,
                    state.FinalPurpose,
                    state.FinalTarget,
                    state.FinalUsageContext,
                    state.FinalCoreValue,
                    state.ConstraintCharacterLimit,
                    state.ConstraintCultural,
                    state.ConstraintLegal,
                    state.ConstraintOther,
                    cancellationToken);
                if (!string.IsNullOrWhiteSpace(gen))
                {
                    Console.WriteLine($"[USER_MESSAGE] {gen}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(gen, gen), cancellationToken);
                    var followMsg = "この要素を使ってキャッチフレーズを作ります。";
                    Console.WriteLine($"[USER_MESSAGE] {followMsg}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(followMsg, followMsg), cancellationToken);
                    state.Step7Completed = true;
                }
            }
    private static void WriteColored(string text, ConsoleColor color)
    {
        var prev = Console.ForegroundColor;
        try
        {
            Console.ForegroundColor = color;
            Console.WriteLine(text);
        }
        finally
        {
            Console.ForegroundColor = prev;
        }
    }

    // Step6 生成テキストをシンプルJSON { "状況":[], "課題・欲求":[], "感情":[], "温度感":[] } に変換
    private static Dictionary<string, List<string>>? ParseCreativeElementsToSimpleJson(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var dict = CreativeElementCategories.All
            .ToDictionary(k => k, _ => new List<string>());

        string? current = null;
        var lines = raw.Replace("\r", string.Empty).Split('\n');
        var categoryMap = CreativeElementCategories.All.ToDictionary(c => c, c => c); // エイリアス無し（互換不要）

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (line.Length == 0) continue; // 空行はスキップ

            // カテゴリ見出し (【...】)
            if (line.StartsWith("【") && line.EndsWith("】"))
            {
                var name = line.Trim('【', '】').Trim();
                if (categoryMap.TryGetValue(name, out var mapped))
                {
                    current = mapped;
                }
                else
                {
                    current = null; // 未知カテゴリは無視
                }
                continue;
            }

            if (current == null) continue;

            // 先頭の箇条書き記号除去
            string content = line;
            if (content.StartsWith("- ")) content = content.Substring(2).Trim();
            else if (content.StartsWith("・")) content = content.Substring(1).Trim();
            else if (content.StartsWith("-")) content = content.Substring(1).Trim();
            if (content.Length == 0) continue;
            dict[current].Add(content);
        }

        // 1カテゴリも埋まっていなければ失敗扱い
        if (dict.All(kv => kv.Value.Count == 0)) return null;
        return dict;
    }

    private static string StripFences(string text)
    {
        var t = text.Trim();
        if (t.StartsWith("```"))
        {
            var nl = t.IndexOf('\n');
            if (nl > 0) t = t.Substring(nl + 1).Trim();
            var last = t.LastIndexOf("```", StringComparison.Ordinal);
            if (last >= 0) t = t.Substring(0, last).Trim();
        }
        return t;
    }

    private static Dictionary<string, List<string>>? TryParseCreativeElementsJson(string raw)
    {
        try
        {
            var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;
            string[] keys = CreativeElementCategories.All; // 中央定義
            var dict = new Dictionary<string, List<string>>();
            foreach (var k in keys)
            {
                if (!root.TryGetProperty(k, out var arr) || arr.ValueKind != JsonValueKind.Array) return null;
                var list = new List<string>();
                foreach (var el in arr.EnumerateArray())
                {
                    if (el.ValueKind == JsonValueKind.String)
                    {
                        var v = el.GetString();
                        if (!string.IsNullOrWhiteSpace(v)) list.Add(v!.Trim());
                    }
                }
                if (list.Count != 5) return null; // 厳格: 必ず5件
                dict[k] = list;
            }
            return dict;
        }
        catch { return null; }
    }

    private static string RenderCreativeElementsForUser(Dictionary<string, List<string>> dict)
    {
        string Format(string label) => dict.TryGetValue(label, out var list) && list.Count > 0
            ? "【" + label + "】\n- " + string.Join("\n- ", list)
            : "";
        var parts = CreativeElementCategories.All
            .Select(Format)
            .Where(s => !string.IsNullOrWhiteSpace(s));
        return string.Join("\n\n", parts);
    }

    // コードフェンスや余計な前後文を含むLLM出力から最初のJSONオブジェクトを抽出
    private static string? ExtractFirstJsonObject(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var text = raw.Trim();
        // ```json などのフェンス除去
        if (text.StartsWith("```"))
        {
            // 先頭フェンス行を削る
            var firstLineEnd = text.IndexOf('\n');
            if (firstLineEnd > 0) text = text.Substring(firstLineEnd + 1).Trim();
            // 末尾 ``` を削る
            var lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (lastFence >= 0) text = text.Substring(0, lastFence).Trim();
        }
        int firstBrace = text.IndexOf('{');
        if (firstBrace < 0) return null;
        int depth = 0; bool inStr = false; bool esc = false;
        for (int i = firstBrace; i < text.Length; i++)
        {
            char c = text[i];
            if (inStr)
            {
                if (esc) { esc = false; }
                else if (c == '\\') esc = true;
                else if (c == '"') inStr = false;
            }
            else
            {
                if (c == '"') inStr = true;
                else if (c == '{') depth++;
                else if (c == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        var slice = text.Substring(firstBrace, i - firstBrace + 1);
                        return slice.Trim();
                    }
                }
            }
        }
        return null;
    }

    // Step8: Excel 出力（OneDrive）
    // 戻り値: true の場合、セッションがリセットされ新しい state に切り替わった（呼び出し側は以降旧 state を使わない）
    private async Task<bool> TryStep8ExcelAsync(ITurnContext turnContext, ElicitationState state, CancellationToken cancellationToken)
    {
        if (_oneDriveExcelService == null)
        {
            var skip = "（Excel出力は未構成のためスキップしました。OneDrive 環境変数を設定すれば自動生成されます。）";
            Console.WriteLine($"[USER_MESSAGE] {skip}");
            await turnContext.SendActivityAsync(MessageFactory.Text(skip, skip), cancellationToken);
            return false;
        }
        try
        {
            string? uploadedUrl = null;
            var progress = new Progress<string>(url =>
            {
                uploadedUrl = url;
                // ユーザー選択メッセージ（候補2）: クロスマトリクス法でこれから順次キャッチフレーズを埋めていく案内
                // NOTE: Teams 等クライアントで URL の自動リンク化が全角コロン（：, U+FF1A）直後だと行われない場合があるため
                // 半角コロン + 改行を使って URL のオートリンクを確実にする。
                // 例: "...確認できます:\nhttps://..."
                var matrixMsg = $"このファイル上でクロスマトリクス法を使い、順次キャッチフレーズを埋めていきます。進捗はこちらから確認できます:\n{url}";
                Console.WriteLine($"[USER_MESSAGE] {matrixMsg}");
                // 進捗案内を送信（同期待機）
                turnContext.SendActivityAsync(MessageFactory.Text(matrixMsg, matrixMsg), cancellationToken).GetAwaiter().GetResult();
            });
            // 常に直前で有効トークンを確保（期限切れならサイレント更新）
            var delegatedToken = await EnsureGraphTokenAsync(turnContext, state, cancellationToken);
            if (string.IsNullOrWhiteSpace(delegatedToken))
            {
                // サインインカードを表示済み（次ターンで再開）
                return false;
            }

            var result = await _oneDriveExcelService.CreateAndFillExcelAsync(progress, state.TaglineSummaryJson, cancellationToken, delegatedToken);
            if (result.IsSuccess && !string.IsNullOrWhiteSpace(result.WebUrl))
            {
                var done = $"Excel出力完了: {result.WebUrl}";
                Console.WriteLine($"[USER_MESSAGE] {done}");
                await turnContext.SendActivityAsync(MessageFactory.Text(done, done), cancellationToken);
                state.Step8Completed = true;
                state.CompletedUtc = DateTimeOffset.UtcNow;
                // 旧セッションの履歴を明示クリア（新セッションへ引き継がない保証を強調）
                state.History.Clear();

                // 完了後ガイダンス（自動で新規セッションを始めるか案内のみ）→ここでは案内+自動新規開始
                var guidance = "このセッションは完了しました。新しい案件を開始します。キャッチコピー作成の目的を一言で教えてください。"; // 将来: カスタマイズ可能
                // 新しいセッションへ差し替え
                var newState = ElicitationState.CreateNew();
                if (_userState != null)
                {
                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, newState, cancellationToken);
                    await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                Console.WriteLine($"[USER_MESSAGE] {guidance}");
                await turnContext.SendActivityAsync(MessageFactory.Text(guidance, guidance), cancellationToken);
                return true; // セッションリセット通知
            }
            else
            {
                var err = $"Excel出力失敗: {result.Error ?? "不明なエラー"}";
                Console.WriteLine($"[USER_MESSAGE] {err}");
                await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
                return false;
            }
        }
        catch (Exception ex)
        {
            var err = $"Excel出力中に例外: {ex.Message}";
            Console.WriteLine($"[USER_MESSAGE] {err}");
            await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
            return false;
        }
    }

    // Aパターン: 毎回利用直前にフレームワークからトークンを取得（サイレント更新）し、無ければ OAuthPrompt を開始
    private async Task<string?> EnsureGraphTokenAsync(ITurnContext turnContext, ElicitationState state, CancellationToken ct)
    {
        var connectionName = Environment.GetEnvironmentVariable("BOT_OAUTH_CONNECTION_NAME") ?? "GraphDelegated";
        // 期限が5分未満なら必ず再取得試行（state に保存された古い値は盲信しない）
        bool nearingExpiry = state.DelegatedTokenExpiresUtc != null && (state.DelegatedTokenExpiresUtc.Value - DateTimeOffset.UtcNow) < TimeSpan.FromMinutes(5);

        if (!nearingExpiry && !string.IsNullOrWhiteSpace(state.DelegatedGraphToken))
        {
            return state.DelegatedGraphToken; // 有効期限に余裕があるのでそのまま利用
        }

        try
        {
            // CloudAdapter 環境では TurnState から UserTokenClient を取得するのが推奨
            object? userTokenClientObj = null;
            foreach (var kv in turnContext.TurnState)
            {
                var t = kv.Value?.GetType()?.FullName;
                if (t == "Microsoft.Bot.Builder.Integration.AspNet.Core.UserTokenClient" || t == "Microsoft.Bot.Connector.Authentication.UserTokenClientImpl")
                {
                    userTokenClientObj = kv.Value; break;
                }
            }
            if (userTokenClientObj != null)
            {
                var method = userTokenClientObj.GetType().GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
                    .FirstOrDefault(m => m.Name == "GetUserTokenAsync" && m.GetParameters().Length >= 4);
                if (method != null)
                {
                    // 代表的シグネチャ (userId, connectionName, channelId, magicCode, cancellationToken)
                    var ps = method.GetParameters();
                    object?[] args;
                    if (ps.Length == 5)
                        args = new object?[] { turnContext.Activity.From?.Id, connectionName, turnContext.Activity.ChannelId, null, ct };
                    else if (ps.Length == 4)
                        args = new object?[] { turnContext.Activity.From?.Id, connectionName, turnContext.Activity.ChannelId, ct };
                    else
                        args = Array.Empty<object?>();
                    if (args.Length > 0)
                    {
                        var taskObj = method.Invoke(userTokenClientObj, args);
                        if (taskObj is Task task)
                        {
                            await task.ConfigureAwait(false);
                            var resultProp = task.GetType().GetProperty("Result");
                            var tokenResponse = resultProp?.GetValue(task);
                            var tokenProp = tokenResponse?.GetType().GetProperty("Token");
                            var tokenVal = tokenProp?.GetValue(tokenResponse) as string;
                            if (!string.IsNullOrWhiteSpace(tokenVal))
                            {
                                state.DelegatedGraphToken = tokenVal;
                                state.LastTokenAcquiredUtc = DateTimeOffset.UtcNow;
                                state.WaitingForSignIn = false;
                                state.DelegatedTokenExpiresUtc = TryReadJwtExpiry(tokenVal);
                                if (_userState != null)
                                {
                                    var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                                    await accessor.SetAsync(turnContext, state, ct);
                                    await _userState.SaveChangesAsync(turnContext, false, ct);
                                }
                                return tokenVal;
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[OAUTH][Ensure] silent acquisition failed: {ex.Message}");
        }

        // ここまでで取得できなかった → OAuthPrompt 起動
        if (_conversationState != null && _mainDialog != null)
        {
            state.WaitingForSignIn = true;
            var guide = "OneDrive へのアクセス許可が必要です。サインインカードを表示しますので認証してください。";
            await turnContext.SendActivityAsync(MessageFactory.Text(guide, guide), ct);
            try
            {
                var dialogStateAccessor = _conversationState.CreateProperty<DialogState>("DialogState");
                Console.WriteLine("[OAUTH][Ensure] MainDialog invoked (token missing)");
                await _mainDialog.RunAsync(turnContext, dialogStateAccessor, ct);
                await _conversationState.SaveChangesAsync(turnContext, false, ct);
            }
            catch (Exception exStart)
            {
                Console.WriteLine("[OAUTH][Ensure][ERROR] " + exStart.Message);
                await turnContext.SendActivityAsync(MessageFactory.Text("サインインカードの起動に失敗しました: " + exStart.Message), ct);
            }
        }
        return null;
    }

    private static DateTimeOffset? TryReadJwtExpiry(string token)
    {
        try
        {
            var handler = new JwtSecurityTokenHandler();
            if (!handler.CanReadToken(token)) return null;
            var jwt = handler.ReadJwtToken(token);
            return jwt.ValidTo == DateTime.MinValue ? null : new DateTimeOffset(jwt.ValidTo, TimeSpan.Zero);
        }
        catch { return null; }
    }

}

// 目的引き出し用の会話状態
public class ElicitationState
{
    public List<string> History { get; } = new List<string>();
    // SSO サインイン待機中フラグ (Excel 出力直前にトークン未取得だった場合 SignInCard 送信し true)
    public bool WaitingForSignIn { get; set; } = false;
    // OAuthPrompt で取得した Graph 委任トークン（ランタイム短期保存。長期保存不要）
    public string? DelegatedGraphToken { get; set; }
    public string? FinalPurpose { get; set; }
    public string SessionId { get; set; } = Guid.NewGuid().ToString();
    public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset? CompletedUtc { get; set; }
    public string? Step1SummaryJson { get; set; }
    // Step1の質問回数（くどさを抑えるためのペース制御）
    public int Step1QuestionCount { get; set; } = 0;
    // ウェルカムメッセージを送ったかどうか（重複送信防止用）
    public bool WelcomeSent { get; set; } = false;
    // フェーズ管理
    public bool Step1Completed { get; set; } = false;
    public bool Step2Completed { get; set; } = false;
    public bool Step3Completed { get; set; } = false;
    public bool Step4Completed { get; set; } = false;
    public bool Step5Completed { get; set; } = false; // 制約事項（文字数/文化/法規/その他）
    public bool Step6Completed { get; set; } = false; // 要約生成（キャッチフレーズ用）
    public bool Step7Completed { get; set; } = false; // クリエイティブ要素自動生成
    public bool Step8Completed { get; set; } = false; // Excel出力
    public string? TaglineSummaryJson { get; set; } // Step6生成結果
    public string? ExcelItemId { get; set; } // Step8 で取得
    // Step2（ターゲット）
    public int Step2QuestionCount { get; set; } = 0;
    public string? FinalTarget { get; set; }
    // Step3（媒体）
    public int Step3QuestionCount { get; set; } = 0;
    public string? FinalUsageContext { get; set; }
    // Step4（コア）
    public int Step4QuestionCount { get; set; } = 0;
    public string? FinalCoreValue { get; set; }
    // Step5（制約事項）
    public int Step5QuestionCount { get; set; } = 0;
    public string? ConstraintCharacterLimit { get; set; }
    public string? ConstraintCultural { get; set; }
    public string? ConstraintLegal { get; set; }
    public string? ConstraintOther { get; set; }
    // Step6（クリエイティブ要素）: 質問カウント不要
    // --- Diagnostics (SSO) ---
    public DateTimeOffset? LastTokenAttemptUtc { get; set; }
    public string? LastTokenResult { get; set; } // success / null / exception-message(short)
    public string? LastDelegatedTokenPreview { get; set; } // masked
    // OAuth diagnostics
    public int OAuthPromptStartCount { get; set; } = 0;
    public DateTimeOffset? OAuthPromptLastAttemptUtc { get; set; }
    public DateTimeOffset? LastTokenAcquiredUtc { get; set; }
    public DateTimeOffset? DelegatedTokenExpiresUtc { get; set; } // JWT exp (診断用)
    // --- Evaluation Reasons (for guiding next questions) ---
    public string? LastPurposeReason { get; set; }
    public string? LastTargetReason { get; set; }
    public string? LastMediaReason { get; set; }
    public string? LastCoreReason { get; set; }
    public string? LastConstraintsReason { get; set; }

    public static ElicitationState CreateNew()
    {
        return new ElicitationState
        {
            SessionId = Guid.NewGuid().ToString(),
            StartedUtc = DateTimeOffset.UtcNow,
            CompletedUtc = null,
            Step1Completed = false,
            Step2Completed = false,
            Step3Completed = false,
            Step4Completed = false,
            Step5Completed = false,
            Step6Completed = false,
            Step7Completed = false,
            Step8Completed = false
        };
    }
}

// 評価者の判定結果
public class EvalDecision
{
    public bool IsSatisfied { get; set; }
    public string? Purpose { get; set; }
    public string? Reason { get; set; }
}

// ターゲット評価結果
public class TargetDecision
{
    public bool IsSatisfied { get; set; }
    public string? Target { get; set; }
    public string? Reason { get; set; }
}

// 媒体/利用シーンの評価結果
public class MediaDecision
{
    public bool IsSatisfied { get; set; }
    public string? MediaOrContext { get; set; }
    public string? Reason { get; set; }
}

// コア（提供価値・差別化）評価結果
public class CoreDecision
{
    public bool IsSatisfied { get; set; }
    public string? Core { get; set; }
    public string? Reason { get; set; }
}

// フロー: Step1（目的）→ Step2（ターゲット）→ Step3（媒体/利用シーン）→ Step4（提供価値・差別化のコア）→ Step5（要素アイデア自動生成）