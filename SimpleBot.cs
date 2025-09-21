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

            // ステップ初期化（既定は Purpose）
            state.CurrentStep ??= Steps.Purpose;

            // 履歴にユーザー発話を追加
            state.History.Add($"User: {text}");
            TrimHistory(state.History, 30);

            // ルーティング：ステップごとに処理
            var transcript = string.Join("\n", state.History);
            if (state.CurrentStep == Steps.Purpose)
            {
                // Step1: 目的評価
                var eval = await EvaluatePurposeAsync(_kernel, transcript, cancellationToken);
                if (eval != null && eval.IsSatisfied)
                {
                    // 目的が十分に引き出せた
                    state.FinalPurpose = string.IsNullOrWhiteSpace(eval.Purpose) ? state.FinalPurpose : eval.Purpose;
                    var purposeText = state.FinalPurpose ?? eval.Purpose ?? "(未取得)";

                    // 要約生成（JSON）→ ユーザー向け整形
                    var summaryJson = await GeneratePurposeSummaryAsync(_kernel, transcript, purposeText, cancellationToken);
                    state.Step1SummaryJson = summaryJson;
                    var summaryView = RenderPurposeSummaryForUser(summaryJson, fallbackPurpose: purposeText);

                    var doneMsg = $"目的が見えてきました。\n\n— ここまでの整理 —\n{summaryView}";
                    Console.WriteLine($"[USER_MESSAGE] {doneMsg}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(doneMsg, doneMsg), cancellationToken);

                    // ステップを Target へ遷移し、最初の質問を提示
                    state.CurrentStep = Steps.Target;
                    var targetAsk = await GenerateNextTargetQuestionAsync(_kernel, transcript, purposeText, cancellationToken);
                    if (string.IsNullOrWhiteSpace(targetAsk))
                    {
                        // 1回だけAIに再試行を委任（テンプレ固定文なし）
                        targetAsk = await GenerateNextTargetQuestionAsync(_kernel, transcript, purposeText, cancellationToken);
                    }

                    if (!string.IsNullOrWhiteSpace(targetAsk))
                    {
                        state.History.Add($"Assistant: {targetAsk}");
                        TrimHistory(state.History, 30);
                        Console.WriteLine($"[USER_MESSAGE] {targetAsk}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(targetAsk, targetAsk), cancellationToken);
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
                // Step1のみサニタイズ適用（ターゲット系禁止）
                ask = SanitizeElicitorQuestion(ask);
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
            else if (state.CurrentStep == Steps.Target)
            {
                // Step2: ターゲット引き出し（初期は質問のみ）
                var purposeText = state.FinalPurpose ?? "(目的未保存)";
                var ask = await GenerateNextTargetQuestionAsync(_kernel, transcript, purposeText, cancellationToken);
                if (string.IsNullOrWhiteSpace(ask))
                {
                    // 1回だけAIに再試行（テンプレ固定文なし）
                    ask = await GenerateNextTargetQuestionAsync(_kernel, transcript, purposeText, cancellationToken);
                }
                if (!string.IsNullOrWhiteSpace(ask))
                {
                    state.History.Add($"Assistant: {ask}");
                    TrimHistory(state.History, 30);

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

            // ここには通常到達しません（すべての分岐で return）。
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
        // 初回接続時のメッセージ
    var welcomeMessage = "こんにちは。今日はどんな言葉づくりをお手伝いしましょう？差し支えなければ、使う場面をひとつ教えてください。（例：ポスター、SNS、Web など）";
        Console.WriteLine($"[USER_MESSAGE] {welcomeMessage}");
        await turnContext.SendActivityAsync(MessageFactory.Text(welcomeMessage, welcomeMessage), cancellationToken);

        await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);
    }

    private static void TrimHistory(List<string> history, int max)
    {
        if (history.Count > max)
        {
            history.RemoveRange(0, history.Count - max);
        }
    }

        private static async Task<EvalDecision?> EvaluatePurposeAsync(Kernel kernel, string transcript, CancellationToken ct)
    {
    var chat = kernel.GetRequiredService<IChatCompletionService>();
    var history = new ChatHistory();
        history.AddSystemMessage(@"あなたは第三者のレビュアーです。今は『目的の明確化』ステップのみを評価します。
ターゲット設定、商品・サービスの本質整理、表現案（キャッチコピー案）作成などは別フェーズです。
これらには触れず、評価にも含めないでください。キャッチコピー作成における『目的』が十分に具体的かを判定してください。

内部定義（ユーザーには明示しない）として、『目的の明確化』は次の観点に十分近づくことを指します：
1) コピーを使う場面（例：広告、Webサイト、SNS投稿、イベント告知、商品パッケージ）
2) コピーの最終的な役割（例：認知を広げたい、購買を促したい、ブランドイメージを浸透させたい）
3) 期待する行動（例：商品を買う、イベントに参加する、サイトを訪れる、記憶に残す）
4) 到達したいゴールを一文で表す（例：『このコピーで、まずブランド名を覚えてもらう』）
5) コアの理由・背景（なぜ今それをするのか／なぜこの手法［例：ポスター］なのか）
6) 表現上の必須事項や避けたいこと（must include / must avoid）
7) タイミング・期限・配置や制作上の制約があれば把握

必要十分性の判定ルール：
- 少なくとも (1)〜(4) のうち複数要素が具体化し、かつ (5) の『コアの理由・背景（なぜ今／なぜこの手法）』に触れていること。
- (6)(7) は任意だが、明確に『未定・不明』と分かる形で扱われたなら許容。完全に触れられていない場合は不足として指摘。
これらを満たさない場合は isSatisfied=false にしてください。

出力は次のJSONのみ。余計なテキストは出さないこと:
{
  ""isSatisfied"": true,
  ""purpose"": ""十分に明確なら短く要約。未確定なら空文字でもよい"",
  ""reason"": ""判断理由を簡潔に（不足している要素があれば指摘）""
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
    history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名や段取りはユーザーに説明せず、
自然な対話で必要な情報を引き出します。テンプレの型を押し付けず、誘導感を避け、許可ベースで尋ねます（例：『もしよければ…』『差し支えなければ…』）。
ユーザーがターゲットや表現案など別の話題に触れても、『それも大切ですね。まず状況だけ少し教えてください——』のようにさりげなく軌道を戻し、
目的に関する短い質問を1つだけ投げかけてください。曖昧な回答は受け止め、選択肢の例を軽く示すのは可（押し付けない）。

質問密度は控えめにしてください{pacing}。『なぜ？』の連鎖や細かい詰め問いは避けます。

内部ルール：ここでは『目的』を整えます。次の観点に近づくほど良いですが、この分類はユーザーに言いません。
1) コピーを使う場面（例：広告、Webサイト、SNS投稿、イベント告知、商品パッケージ）
2) コピーの最終的な役割（例：認知を広げたい、購買を促したい、ブランドイメージを浸透させたい）
3) 期待する行動（例：商品を買う、イベントに参加する、サイトを訪れる、記憶に残す）
4) 到達したいゴールを一文で表す（例：『このコピーで、まずブランド名を覚えてもらう』）
5) コアの理由・背景（なぜ今それをするのか／なぜこの手法［例：ポスター］なのか）
6) 表現上の必須事項・避けたいこと（must include / must avoid）
7) タイミング・期限・配置や制作上の制約（あれば）

以下の会話を踏まえて、目的をはっきりさせるために、短く1つだけ質問してください。専門用語は避け、
日常の言い回しで。必要に応じて“仮の一文”や“方向性の仮説を2〜3個（各1行）”を示し、『こんな感じでしょうか？』と軽く確認してもOKです（断定しない）。

禁止・NG（質問に含めない）:
- 「誰に」「どんな人」「ターゲット」「対象」「◯◯向け」「年齢」「性別」「属性」「居住地域」「職業」「層」「ペルソナ」など、
  人物像・属性・セグメントに関する質問（これらは後続ステップで扱います。今は目的だけに集中してください。）

質問のヒント（必要に応じて1つ選んで質問）：
- どんな場面でこの言葉を使いますか？（例：ポスター、SNS、会場アナウンス など）
- このコピーは最終的にどんな働きをしてほしいですか？（例：まず知ってもらう／行動を後押しする など）
- 見たあとに、どんな行動をしてほしいですか？（ざっくりでOK）
- 一言で言うと、何ができたら成功ですか？（一文で）
- 差し支えなければ、今回は『なぜ今』取り組むのか教えてもらえますか？（背景・きっかけ）
- もしポスターを選んだ理由があれば教えてください（他の手段ではなくポスターにした訳）
- これだけは入れたい／避けたい表現はありますか？（例：NGワード、トーンの注意）
- 時期や掲出場所・サイズなど、決まっている条件はありますか？（ざっくりでOK）

 出力はユーザーにそのまま見せる日本語テキストだけ。前置きやルール説明、工程名は書かないでください。質問は1つ。許可ベース・選択肢例はOK。
");
        history.AddUserMessage($"--- 会話履歴 ---\n{transcript}");

        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
    }

    // Step1（目的）サマリーの生成（JSON）
        private static async Task<string> GeneratePurposeSummaryAsync(Kernel kernel, string transcript, string purpose, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
                                history.AddSystemMessage(@"あなたは第三者のレビュアー兼まとめ役です。今は『目的の明確化』ステップの成果を、
ユーザーに見せる要約（JSON）にまとめます。事実のみを使い、推測や創作はしません。ユーザーに分類を意識させる必要はありません。

内部定義：『目的の明確化』は次の観点を指します（欠けていても可）。
1) 使用場面 2) 最終的な役割 3) 期待する行動 4) 到達ゴール（一文） 5) コアの理由・背景 6) 必須/NG 7) タイミング・制約

もし他ステップ（ターゲット/本質/表現案）に関わる情報が会話に出ていれば、reference として別枠に記録します。
出力は次のJSONのみ：
{
  ""purpose"": ""最終的な目的の要約（1行）"",
  ""usageContext"": ""判明していれば1行（なければ空）"",
  ""finalRole"": ""判明していれば1行（なければ空）"",
  ""expectedAction"": ""判明していれば1行（なければ空）"",
  ""oneSentenceGoal"": ""可能なら一文（なければ空）"",
    ""coreDriver"": ""なぜ今／なぜこの手法（例：ポスター）なのか。背景やきっかけ（なければ空）"",
    ""mustInclude"": [""入れたい要素（なければ空配列）""],
    ""mustAvoid"": [""避けたい要素（なければ空配列）""],
    ""timingOrConstraints"": ""時期・期限・掲出条件・制作上の制約など（なければ空）"",
  ""references"": {
    ""targetHints"": [""出ていれば箇条書き""],
    ""essenceHints"": [""出ていれば箇条書き""],
    ""expressionHints"": [""出ていれば箇条書き""]
  },
  ""gaps"": [""不足している点（なければ空配列）""],
    ""confidence"": 0.0,
    ""reviewer"": {
        ""evaluation"": ""目的の充足度に関する短い所見（ユーザーに直接は見せない）"",
        ""missingPoints"": [""不足/曖昧な点（例：使用場面が不明確）""],
        ""discomfortSignals"": [""ユーザーが不快に感じていそうな兆し（あれば）""],
        ""guidanceForNextAI"": ""次のAI（質問をする側）への注意や配慮（ユーザーには見せない）""
    }
}
");
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
            if (!string.IsNullOrWhiteSpace(usage)) lines.Add($"使用場面: {usage}");
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

    // Step2（ターゲット）質問生成
    private static async Task<string?> GenerateNextTargetQuestionAsync(Kernel kernel, string transcript, string purpose, CancellationToken ct)
    {
        var chat = kernel.GetRequiredService<IChatCompletionService>();
        var history = new ChatHistory();
        // レビュアーの所見（あれば）を次のAIプロンプトに注入する準備
        string? reviewerNote = ExtractReviewerNoteFromTranscript(transcript);
    var systemHeader = @"あなたはやさしく話すマーケティングのプロです。工程名はユーザーに見せず、
自然な対話で“相手像”を絞る手助けをします。テンプレの型を押し付けず、誘導感を避け、許可ベースで尋ねます。
ユーザーが別の話題に触れても、『ありがとうございます。相手像を少しだけ教えてください——』のようにやわらかく戻し、短い質問を1つだけしてください。

もしユーザーが『くどい』『細かい』などのサインを出したら、質問は控えめにして、仮の相手像を1〜3行で提案し『ざっくりこんなイメージで大丈夫？（違っていれば一言だけ補足ください）』とだけ確認してください。

出力はユーザーにそのまま見せる日本語テキストのみ。前置きやルール説明は書かないこと。質問は1つ。許可ベース・例示はOK。

聞き方の例：
- どんな人に届くと良さそうですか？
- どんな特徴の人を想定していますか？（例：近隣に住む家族連れ、学生、地域のシニア など）
- 今回は特に来てほしい人はいますか？（ざっくりでOK）
";
        if (!string.IsNullOrWhiteSpace(reviewerNote))
        {
            systemHeader += "\n\n[内部メモ（ユーザーに見せない）]\n" + reviewerNote + "\n";
        }
        history.AddSystemMessage(systemHeader);
        history.AddUserMessage($"— 目的 —\n{purpose}\n\n— 会話履歴 —\n{transcript}");
        var response = await chat.GetChatMessageContentAsync(history, kernel: kernel, cancellationToken: ct);
        return response?.Content?.Trim();
    }

    // transcript内の直近の要約JSONから reviewer を抽出し、次AI用の注意メモを生成
    private static string? ExtractReviewerNoteFromTranscript(string transcript)
    {
        // 直近期のJSONブロックを雑に抽出して解析（厳密なパースは不要。例外安全に）
        try
        {
            // 最後に現れる '{' から終端 '}' までを拾ってJSONとして試験的に読む
            int start = transcript.LastIndexOf('{');
            int end = transcript.LastIndexOf('}');
            if (start >= 0 && end > start)
            {
                var json = transcript.Substring(start, end - start + 1);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (root.TryGetProperty("reviewer", out var rev) && rev.ValueKind == JsonValueKind.Object)
                {
                    var eval = rev.TryGetProperty("evaluation", out var e) && e.ValueKind == JsonValueKind.String ? e.GetString() : null;
                    var guidance = rev.TryGetProperty("guidanceForNextAI", out var g) && g.ValueKind == JsonValueKind.String ? g.GetString() : null;
                    var missing = rev.TryGetProperty("missingPoints", out var m) && m.ValueKind == JsonValueKind.Array ? string.Join("; ", m.EnumerateArray().Select(x => x.GetString()).Where(s => !string.IsNullOrWhiteSpace(s))) : null;
                    var discomfort = rev.TryGetProperty("discomfortSignals", out var d) && d.ValueKind == JsonValueKind.Array ? string.Join("; ", d.EnumerateArray().Select(x => x.GetString()).Where(s => !string.IsNullOrWhiteSpace(s))) : null;

                    var parts = new List<string>();
                    if (!string.IsNullOrWhiteSpace(eval)) parts.Add($"所見: {eval}");
                    if (!string.IsNullOrWhiteSpace(missing)) parts.Add($"不足/曖昧: {missing}");
                    if (!string.IsNullOrWhiteSpace(discomfort)) parts.Add($"不快の兆し: {discomfort}");
                    if (!string.IsNullOrWhiteSpace(guidance)) parts.Add($"次AIへの配慮: {guidance}");
                    return parts.Count > 0 ? string.Join("\n", parts) : null;
                }
            }
        }
        catch { /* 失敗しても無視（注入しない） */ }
        return null;
    }

    // Elicitorの出力をスコープ遵守に補正するガード
    private static string SanitizeElicitorQuestion(string? ask)
    {
        var s = (ask ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(s))
        {
            return "このポスターで何が起きたら『うまくいった』と言えそうですか？（ざっくりでOK）";
        }

        string[] banned = new[]
        {
            "誰に","どんな人","ターゲット","対象","向け","年齢","性別","属性","居住","地域の人","職業","層","ペルソナ"
        };

        foreach (var b in banned)
        {
            if (s.Contains(b, StringComparison.OrdinalIgnoreCase))
            {
                // ターゲット系に触れていれば安全な質問に差し替え
                return "どんな場面でこの言葉を使いますか？（例：ポスター、SNS、会場アナウンス など）";
            }
        }

        // 長すぎる出力は切り詰め（安全側）
        if (s.Length > 120) s = s.Substring(0, 120);
        return s;
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
    public Steps? CurrentStep { get; set; }
    public string? Step1SummaryJson { get; set; }
    // Step1の質問回数（くどさを抑えるためのペース制御）
    public int Step1QuestionCount { get; set; } = 0;
}

// 評価者の判定結果
public class EvalDecision
{
    public bool IsSatisfied { get; set; }
    public string? Purpose { get; set; }
    public string? Reason { get; set; }
}

public enum Steps
{
    Purpose,
    Target,
    Essence,
    Expression
}