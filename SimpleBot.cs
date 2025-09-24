using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.Extensions.DependencyInjection;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;
using System;

public class SimpleBot : ActivityHandler
{
    private readonly Kernel? _kernel; // optional
    private readonly ConversationState? _conversationState;

    public SimpleBot(Kernel? kernel = null, ConversationState? conversationState = null)
    {
        _kernel = kernel;
        _conversationState = conversationState;
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
            var kernelErrorMessage = "Kernel が初期化されていません。";
            Console.WriteLine($"[USER_MESSAGE] {kernelErrorMessage}");
            await turnContext.SendActivityAsync(MessageFactory.Text(kernelErrorMessage, kernelErrorMessage), cancellationToken);
            return;
        }

        try
        {
            // 会話状態取得（ユーザー/アシスタントの履歴を保持）
            var state = new ElicitationState();
            if (_conversationState != null)
            {
                var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                state = await accessor.GetAsync(turnContext, () => new ElicitationState(), cancellationToken);
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
                if (tEval != null && tEval.IsSatisfied)
                {
                    state.FinalTarget = string.IsNullOrWhiteSpace(tEval.Target) ? state.FinalTarget : tEval.Target;

                    // 簡潔に承知のみ（Step2はAI生成の承知文は使わず最小限の固定文）
                    var tAck = $"ターゲット像、承知しました。ありがとうございます。\n- {state.FinalTarget}";
                    Console.WriteLine($"[USER_MESSAGE] {tAck}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(tAck, tAck), cancellationToken);

                    state.Step2Completed = true; // ターゲット完了

                    // Step3（媒体/利用シーン）へ最初の短い質問を投げる
                    var mFirst = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, cancellationToken);
                    if (string.IsNullOrWhiteSpace(mFirst))
                    {
                        mFirst = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, cancellationToken);
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
                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 既出の手がかりを踏まえて次のターゲット質問を生成
                var tAsk = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, cancellationToken);

                if (string.IsNullOrWhiteSpace(tAsk))
                {
                    // 1回だけAIに再試行
                    tAsk = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, cancellationToken);
                }

                if (!string.IsNullOrWhiteSpace(tAsk))
                {
                    state.History.Add($"Assistant: {tAsk}");
                    TrimHistory(state.History, 30);
                    state.Step2QuestionCount++;

                    // 状態保存
                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
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
                if (mEval != null && mEval.IsSatisfied)
                {
                    state.FinalUsageContext = string.IsNullOrWhiteSpace(mEval.MediaOrContext) ? state.FinalUsageContext : mEval.MediaOrContext;

                    var mAck = $"媒体／利用シーン、承知しました。ありがとうございます。\n- {state.FinalUsageContext}";
                    Console.WriteLine($"[USER_MESSAGE] {mAck}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(mAck, mAck), cancellationToken);

                    state.Step3Completed = true; // 媒体完了

                    // 状態保存
                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 既出の usageContext を活かして次の質問を生成
                var mAsk = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, cancellationToken);
                if (string.IsNullOrWhiteSpace(mAsk))
                {
                    mAsk = await GenerateNextMediaQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step3QuestionCount, cancellationToken);
                }
                if (!string.IsNullOrWhiteSpace(mAsk))
                {
                    state.History.Add($"Assistant: {mAsk}");
                    TrimHistory(state.History, 30);
                    state.Step3QuestionCount++;

                    // 状態保存
                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
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
                if (cEval != null && cEval.IsSatisfied)
                {
                    state.FinalCoreValue = string.IsNullOrWhiteSpace(cEval.Core) ? state.FinalCoreValue : cEval.Core;

                    var cAck = $"コア（提供価値・差別化）、承知しました。ありがとうございます。\n- {state.FinalCoreValue}";
                    Console.WriteLine($"[USER_MESSAGE] {cAck}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(cAck, cAck), cancellationToken);

                    state.Step4Completed = true; // Step4完了

                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
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
                        cancellationToken);
                }

                if (!string.IsNullOrWhiteSpace(cAsk))
                {
                    state.History.Add($"Assistant: {cAsk}");
                    TrimHistory(state.History, 30);
                    state.Step4QuestionCount++;

                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }

                    Console.WriteLine($"[USER_MESSAGE] {cAsk}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(cAsk, cAsk), cancellationToken);
                }
                return;
            }

            // Step1: 目的評価
            var eval = await EvaluatePurposeAsync(_kernel, transcript, cancellationToken);
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

                var summaryMsg = $"— ここまでの整理 —\n{summaryView}";
                Console.WriteLine($"[USER_MESSAGE] {summaryMsg}");
                await turnContext.SendActivityAsync(MessageFactory.Text(summaryMsg, summaryMsg), cancellationToken);

                // Step1完了フラグ
                state.Step1Completed = true;

                // 直後にStep2（ターゲット）への最初の短い質問を1つだけ行う
                var tFirst = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, cancellationToken);
                if (string.IsNullOrWhiteSpace(tFirst))
                {
                    tFirst = await GenerateNextTargetQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.Step2QuestionCount, cancellationToken);
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
                if (_conversationState != null)
                {
                    var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                }
                return;
            }

            // 未確定: Elicitor に次の質問を生成（Step1用）
            var ask = await GenerateNextQuestionAsync(_kernel, transcript, state.Step1QuestionCount, cancellationToken);

            if (string.IsNullOrWhiteSpace(ask))
            {
                // 1回だけAIに再試行（テンプレ固定文なし）
                ask = await GenerateNextQuestionAsync(_kernel, transcript, state.Step1QuestionCount, cancellationToken);
            }

            if (!string.IsNullOrWhiteSpace(ask))
            {
                state.History.Add($"Assistant: {ask}");
                TrimHistory(state.History, 30);
                // 質問回数をカウントして過度な深掘りを避ける
                state.Step1QuestionCount++;

                // 状態保存
                if (_conversationState != null)
                {
                    var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
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
        if (userJoined)
        {
            var welcomeMessage = "こんにちは。今日はどんな言葉づくりをお手伝いしましょう？まずは、差し支えなければ『何の活動のためのコピーか』を教えてください。（例：イベント告知／販促キャンペーン／ブランド認知／採用 など）";

            if (_conversationState != null)
            {
                var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                var state = await accessor.GetAsync(turnContext, () => new ElicitationState(), cancellationToken);
                if (!state.WelcomeSent)
                {
                    Console.WriteLine($"[USER_MESSAGE] {welcomeMessage}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeMessage, welcomeMessage), cancellationToken);

                    state.WelcomeSent = true;
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
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

    // 目的受領時の承知メッセージをAIに1文で生成させる
    private static async Task<string?> GenerateAckMessageAsync(Kernel kernel, string transcript, string purpose, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは丁寧で軽やかなアシスタントです。以下の会話と受け取った活動目的を踏まえ、
承知の意を短く1文だけ日本語で伝えてください。敬語は自然体で、堅すぎず、命令形や謝罪は避けます。
禁止：目的の言い換えの羅列、評価的コメント、次の質問。
出力はそのままユーザーに見せる1文のみ。");
        history.AddUserMessage($"--- 受領目的 ---\n{purpose}\n\n--- 会話履歴 ---\n{transcript}");
        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
    }

    // Step2: ターゲットが十分に定義されたかを評価
    private static async Task<TargetDecision?> EvaluateTargetAsync(Kernel kernel, string transcript, string? step1SummaryJson, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ2『ターゲット』のみを評価します。
            ここでいうターゲットは、誰に届けるかの像（例：属性、役割、状況、顧客段階など）です。
            会話履歴と、もしあればステップ1要約JSON内の 'audience' や 'references.targetHints' を手掛かりに、既出の情報のみから判断します。

            判定基準：
            - 誰に向けたコピーかが1行で説明できること（例：関東圏の大学1〜2年生、既存のライトユーザー など）
            - 既出の事実に基づくこと（推測や新規追加はしない）

            出力は次のJSONのみ：
            {
                ""isSatisfied"": true,
                ""target"": ""1行の要約（未確定なら空）"",
                ""reason"": ""判断理由（不足点も簡潔に）""
            }");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        var json = response?.Content?.Trim();
        if (string.IsNullOrWhiteSpace(json)) return null;
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True;
            string? target = root.TryGetProperty("target", out var tEl) && tEl.ValueKind == JsonValueKind.String ? tEl.GetString() : null;
            string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null;
            return new TargetDecision { IsSatisfied = isSat, Target = target, Reason = reason };
        }
        catch
        {
            return null;
        }
    }

    // Step2: 次のターゲット確認質問を生成（既出の手がかりを活かす）
    private static async Task<string?> GenerateNextTargetQuestionAsync(Kernel kernel, string transcript, string? step1SummaryJson, int questionCount, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        var pacing = questionCount >= 2
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の選択肢を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『ターゲット』だけを確かめます。
            既出の手がかり（Step1要約のaudienceやtargetHints、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            ターゲット以外（目的の再評価、表現案、制作条件など）は扱いません（別フェーズ）。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            出力はユーザーにそのまま見せる日本語のテキストのみ。");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON（参考） ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
    }

    // Step3: 媒体/利用シーンが十分に定義されたかを評価
    private static async Task<MediaDecision?> EvaluateMediaAsync(Kernel kernel, string transcript, string? step1SummaryJson, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ3『媒体／利用シーン』のみを評価します。
            ここでいう媒体／利用シーンは、キャッチコピーがどこで・どのように使われるか（例：LPのヒーロー、駅ポスター、アプリ内バナー、メール件名 等）です。
            会話履歴と、もしあればステップ1要約JSON内の 'usageContext' を手掛かりに、既出の情報のみから判断します。

            判定基準：
            - 媒体／利用シーンが1行で説明できること（例：特設LPのファーストビュー、店頭A1ポスター など）
            - 既出の事実に基づくこと（推測や新規追加はしない）

            出力は次のJSONのみ：
            {
                ""isSatisfied"": true,
                ""mediaOrContext"": ""1行の要約（未確定なら空）"",
                ""reason"": ""判断理由（不足点も簡潔に）""
            }");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        var json = response?.Content?.Trim();
        if (string.IsNullOrWhiteSpace(json)) return null;
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True;
            string? media = root.TryGetProperty("mediaOrContext", out var mEl) && mEl.ValueKind == JsonValueKind.String ? mEl.GetString() : null;
            string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null;
            return new MediaDecision { IsSatisfied = isSat, MediaOrContext = media, Reason = reason };
        }
        catch
        {
            return null;
        }
    }

    // Step3: 次の媒体確認質問を生成（既出の手がかりを活かす）
    private static async Task<string?> GenerateNextMediaQuestionAsync(Kernel kernel, string transcript, string? step1SummaryJson, int questionCount, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        var pacing = questionCount >= 2
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の選択肢を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『媒体／利用シーン』だけを確かめます。
            既出の手がかり（Step1要約のusageContext、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            媒体以外（目的やターゲットの再確認、表現案、制作条件など）は扱いません（別フェーズ）。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            出力はユーザーにそのまま見せる日本語のテキストのみ。");
        var contextBlock = string.IsNullOrWhiteSpace(step1SummaryJson)
            ? "(Step1要約なし)"
            : step1SummaryJson;
        history.AddUserMessage($"--- Step1要約JSON（参考） ---\n{contextBlock}\n\n--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
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
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ4『提供価値・差別化のコア』のみを評価します。
            ここでいうコアとは、ユーザーにとっての価値の中核や、競合と比べたときの差別化要素です。
            会話履歴と、もしあればStep1要約JSON内の 'coreDriver' や 'subjectOrHero'、'references.essenceHints' を手掛かりに、既出の情報のみから判断します。

            判定基準：
            - どんな価値／差別化を打ち出すのかが1行で説明できること（例：初心者でも10分で設定完了、地元の実例ストーリーで信頼感、等）
            - 既出の事実に基づくこと（推測や新規追加はしない）

            出力は次のJSONのみ：
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

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        var json = response?.Content?.Trim();
        if (string.IsNullOrWhiteSpace(json)) return null;
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True;
            string? core = root.TryGetProperty("core", out var cEl) && cEl.ValueKind == JsonValueKind.String ? cEl.GetString() : null;
            string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null;
            return new CoreDecision { IsSatisfied = isSat, Core = core, Reason = reason };
        }
        catch
        {
            return null;
        }
    }

    // Step4: 次のコア確認質問を生成（既出の手がかりを活かす）
    private static async Task<string?> GenerateNextCoreQuestionAsync(
        Kernel kernel,
        string transcript,
        string? step1SummaryJson,
        string? finalTarget,
        string? finalUsageContext,
        int questionCount,
        CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        var pacing = questionCount >= 2
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の仮説候補を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『提供価値・差別化のコア』だけを確かめます。
            既出の手がかり（Step1要約のcoreDriver/subjectOrHero/essenceHints、確定済みのターゲットや媒体、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            コア以外（目的/ターゲット/媒体の再確認、表現案、制作条件など）は扱いません（別フェーズ）。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            出力はユーザーにそのまま見せる日本語のテキストのみ。");

        var ctx = new List<string>();
        ctx.Add(string.IsNullOrWhiteSpace(step1SummaryJson) ? "(Step1要約なし)" : step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[FinalTarget] {finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[FinalUsageContext] {finalUsageContext}");
        history.AddUserMessage($"--- コンテキスト（参考） ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
    }


        private static async Task<EvalDecision?> EvaluatePurposeAsync(Kernel kernel, string transcript, CancellationToken ct)
    {
    var chat = kernel.GetRequiredService<IChatCompletionService>();
    var history = new ChatHistory();
                history.AddSystemMessage(@"あなたは第三者のレビュアーです。今はステップ1『活動目的（なぜ作るのか）』のみを評価します。
                ターゲット・表現案・商品本質などは扱いません。キャッチコピーが『どんな活動のために作られるのか』だけを見ます。

                活動目的の例：
                - イベントを告知したい
                - キャンペーンで売上を伸ばしたい（販促）
                - ブランド認知を広げたい
                - 採用活動で人を集めたい
                - その他（上記以外）

                判定基準：
                - 上記いずれか（または近いもの）が明確に述べられていること
                - 活動の意図が短い一文で説明できていること
                満たさない場合は isSatisfied=false にしてください。

                出力は次のJSONのみ。余計なテキストは出さないこと:
                {
                    ""isSatisfied"": true,
                    ""purpose"": ""活動目的を短く一文で。未確定なら空でもよい"",
                    ""reason"": ""判断理由を簡潔に（不足があれば指摘）""
                }");
        history.AddUserMessage($"--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        var json = response?.Content?.Trim();
        if (string.IsNullOrWhiteSpace(json)) return null;
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            bool isSat = root.TryGetProperty("isSatisfied", out var satEl) && satEl.ValueKind == JsonValueKind.True;
            string? purpose = root.TryGetProperty("purpose", out var pEl) && pEl.ValueKind == JsonValueKind.String ? pEl.GetString() : null;
            string? reason = root.TryGetProperty("reason", out var rEl) && rEl.ValueKind == JsonValueKind.String ? rEl.GetString() : null;
            return new EvalDecision { IsSatisfied = isSat, Purpose = purpose, Reason = reason };
        }
        catch
        {
            return null;
        }
    }

    private static async Task<string?> GenerateNextQuestionAsync(Kernel kernel, string transcript, int questionCount, CancellationToken ct)
    {
    var chat = kernel.GetRequiredService<IChatCompletionService>();
    var history = new ChatHistory();
    var pacing = questionCount >= 2
        ? "（すでに質問が続いているため、深掘りは控えめに。2〜3個の方向性の仮説を各1行で提案し、『どれが近い？／他にありますか？』とだけ確認。畳みかけはNG）"
        : string.Empty;
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『活動目的』だけを確かめます。
            テンプレ口調は避け、許可ベースで、短く1つだけ質問してください。質問密度は控えめにしてください{pacing}。

            ここで整えるのは『何の活動のためのコピーか』だけです（例：イベント告知／販促キャンペーン／ブランド認知／採用 など）。
            ターゲット像や表現案、制作条件などには触れません（別フェーズ）。

            以下の会話を踏まえて、活動目的をはっきりさせるための短い質問を1つだけ行ってください。必要なら2〜3個の候補（各1行）を示し、
            『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            出力はユーザーにそのまま見せる日本語のテキストのみ。");
        history.AddUserMessage($"--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
    }

    // Step1（目的）サマリーの生成（JSON）
        private static async Task<string> GeneratePurposeSummaryAsync(Kernel kernel, string transcript, string purpose, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
                var history = new ChatHistory();
                history.AddSystemMessage(@"あなたは第三者のレビュアー兼まとめ役です。今はステップ1『活動目的』を中心に、会話中に『既に出ている情報』だけを静かに拾って要約（JSON）に整えます。
                    重要: ユーザーに追加質問はしません。推測や創作は厳禁。会話に現れていない項目は空/空配列にしてください。

                    特に、会話中に出ていれば次を拾います：
                    - どんな媒体/利用シーン（例: ランディングページ、ポスター 等）
                    - 対象・主役（例: 製品名、サービス、イベント名 など）
                    - 大枠ターゲット（例: 学生、既存顧客、来場見込み者 など）
                    - 達成したい効果の種類（ゴールイメージ）

                    出力は次のJSONのみ：
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

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim() ?? "{}";
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
}

public class CatchphraseSkill
{
    [KernelFunction]
    public string DefinePurpose(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return "入力が空です。キャッチコピーの目的を入力してください。";
        }

        // 目的を生成するテンプレート
        return $"キャッチコピーの目的を明確にしました: 「{input}」を基に、ターゲットに響く魅力的なキャッチコピーを作成します。例: 『{input}で世界を変える』";
    }

    [KernelFunction]
    public string SetTarget(string input)
    {
        return $"ターゲットを設定しました: {input}";
    }

    [KernelFunction]
    public string ClarifyEssence(string input)
    {
        return $"商品・サービスの本質を整理しました: {input}";
    }

    // 内部評価用のヘルパー（SKへ直接プロンプト送信）
    public async Task<string> EvaluatePurposeAsync(string input, Kernel kernel)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return "入力が空です。キャッチコピーの目的を入力してください。";
        }

        // 直接プロンプトで評価（グローバル関数登録なし・重複回避）
        var prompt = @"あなたはマーケティングのプロです。以下のユーザー入力が
            キャッチコピー作成の『目的』として十分に明確かを評価してください。
            次のJSONで返答してください（余計なテキストは出力しない）：
            {
                ""isSatisfied"": true,
                ""reason"": ""短い説明""
            }

            ユーザー入力: {{$input}}";

        var arguments = new KernelArguments
        {
            { "input", input }
        };

        var jsonResponse = await kernel.InvokePromptAsync(prompt, arguments);

        // Null チェック
        var jsonString = jsonResponse?.GetValue<string>();
        if (string.IsNullOrEmpty(jsonString))
        {
            return "AIからの応答が無効です。もう一度お試しください。";
        }

        // AIの判断結果を解析
        var result = JsonSerializer.Deserialize<JsonElement>(jsonString);
        if (result.GetProperty("isSatisfied").GetBoolean())
        {
            return "目的が明確になりました。次のステップに進みます。";
        }
        else
        {
            return "目的がまだ明確ではありません。もう少し具体的に入力してください。";
        }
    }
}

// 目的引き出し用の会話状態
public class ElicitationState
{
    public List<string> History { get; } = new List<string>();
    public string? FinalPurpose { get; set; }
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
    // Step2（ターゲット）
    public int Step2QuestionCount { get; set; } = 0;
    public string? FinalTarget { get; set; }
    // Step3（媒体）
    public int Step3QuestionCount { get; set; } = 0;
    public string? FinalUsageContext { get; set; }
    // Step4（コア）
    public int Step4QuestionCount { get; set; } = 0;
    public string? FinalCoreValue { get; set; }
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

// フロー: Step1（目的）→ Step2（ターゲット）→ Step3（媒体/利用シーン）→ Step4（提供価値・差別化のコア）