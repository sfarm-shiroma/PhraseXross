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

public class SimpleBot : ActivityHandler
{
    private readonly Kernel? _kernel; // optional
    private readonly ConversationState? _conversationState;
    private readonly OneDriveExcelService? _oneDriveExcelService; // optional (OneDrive 連携)

    public SimpleBot(Kernel? kernel = null, ConversationState? conversationState = null, OneDriveExcelService? oneDriveExcelService = null)
    {
        _kernel = kernel;
        _conversationState = conversationState;
        _oneDriveExcelService = oneDriveExcelService;
    }

    protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        var text = turnContext.Activity.Text ?? string.Empty;
        Console.WriteLine($"[USER_INPUT] {text}");

        // まず特定コマンドのハンドリング（他フェーズより優先）
        if (text.Trim().Equals("/excel", StringComparison.OrdinalIgnoreCase) || text.Trim().Equals("excel", StringComparison.OrdinalIgnoreCase))
        {
            if (_oneDriveExcelService == null)
            {
                var msg = "OneDrive連携サービスが利用できません（未構成または無効）。";
                Console.WriteLine($"[USER_MESSAGE] {msg}");
                await turnContext.SendActivityAsync(MessageFactory.Text(msg, msg), cancellationToken);
                return;
            }
            try
            {
                var pre = "OneDriveにExcelを作成しています。";
                Console.WriteLine($"[USER_MESSAGE] {pre}");
                await turnContext.SendActivityAsync(MessageFactory.Text(pre, pre), cancellationToken);

                string? uploadedUrl = null;
                var progress = new Progress<string>(url =>
                {
                    uploadedUrl = url;
                    var upMsg = $"アップロード完了（これから内容を書き込みます）: {url}";
                    Console.WriteLine($"[USER_MESSAGE] {upMsg}");
                    turnContext.SendActivityAsync(MessageFactory.Text(upMsg, upMsg), cancellationToken).GetAwaiter().GetResult();
                });

                var result = await _oneDriveExcelService.CreateAndFillExcelAsync(progress, cancellationToken);
                if (result.IsSuccess && !string.IsNullOrWhiteSpace(result.WebUrl))
                {
                    var done = $"Excel出力完了: {result.WebUrl}";
                    Console.WriteLine($"[USER_MESSAGE] {done}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(done, done), cancellationToken);
                }
                else
                {
                    var err = $"Excel出力失敗: {result.Error ?? "不明なエラー"}";
                    Console.WriteLine($"[USER_MESSAGE] {err}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
                }
            }
            catch (Exception exUp)
            {
                var err = $"アップロード処理中に例外: {exUp.Message}";
                Console.WriteLine($"[USER_MESSAGE] {err}");
                await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
            }
            return; // コマンド処理で終了
        }

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

                    var consolidatedAfterStep2 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep2}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(consolidatedAfterStep2, consolidatedAfterStep2), cancellationToken);

                    // Step3（媒体/利用シーン）へ最初の短い質問を投げる（元の自動遷移ロジックを復元）
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

                    var consolidatedAfterStep3 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep3}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(consolidatedAfterStep3, consolidatedAfterStep3), cancellationToken);

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
                        // バリデーション（了承のみで質問しない応答を排除）
                        if (IsInvalidCoreQuestion(coreFirst))
                        {
                            coreFirst = BuildDynamicCoreFallbackQuestion(transcript);
                        }
                        if (string.IsNullOrWhiteSpace(coreFirst))
                        {
                            // フォールバック（AIが空応答だった場合でもユーザーを前に進ませる）
                            coreFirst = "次に『提供価値・差別化のコア』を一言で教えていただけますか？例：地域ならではの温かさ／誰でもすぐ参加できる気軽さ など。";
                            Console.WriteLine("[DEBUG] Core initial question fallback used (AI empty response)");
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

                    var consolidatedAfterStep4 = BuildConsolidatedSummary(state);
                    Console.WriteLine($"[USER_MESSAGE] {consolidatedAfterStep4}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(consolidatedAfterStep4, consolidatedAfterStep4), cancellationToken);

                    // 新Step5（制約事項ヒアリング）へ最初の質問を投げる
                    if (!state.Step5Completed)
                    {
                        var firstConstraintQ = GenerateInitialConstraintQuestion();
                        state.History.Add($"Assistant: {firstConstraintQ}");
                        TrimHistory(state.History, 30);
                        state.Step5QuestionCount++;
                        Console.WriteLine($"[USER_MESSAGE] {firstConstraintQ}");
                        await turnContext.SendActivityAsync(MessageFactory.Text(firstConstraintQ, firstConstraintQ), cancellationToken);
                    }

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

                if (IsInvalidCoreQuestion(cAsk))
                {
                    cAsk = BuildDynamicCoreFallbackQuestion(transcript);
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

            // Step5（制約事項：文字数 / 文化・配慮 / 法規・レギュレーション / その他）フロー
            if (state.Step1Completed && state.Step2Completed && state.Step3Completed && state.Step4Completed && !state.Step5Completed)
            {
                var consEval = await EvaluateConstraintsAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, state.FinalCoreValue, cancellationToken);
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
                    await turnContext.SendActivityAsync(MessageFactory.Text(consolidatedAfterStep5, consolidatedAfterStep5), cancellationToken);

                    // 直後に Step6（旧Step5）クリエイティブ要素自動生成を実行
                    if (!state.Step6Completed)
                    {
                        var step6 = await GenerateCreativeElementsAsync(
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
                        if (!string.IsNullOrWhiteSpace(step6))
                        {
                            Console.WriteLine($"[USER_MESSAGE] {step6}");
                            await turnContext.SendActivityAsync(MessageFactory.Text(step6, step6), cancellationToken);
                            var followMsg = "この要素を使ってキャッチフレーズを作ります。";
                            Console.WriteLine($"[USER_MESSAGE] {followMsg}");
                            await turnContext.SendActivityAsync(MessageFactory.Text(followMsg, followMsg), cancellationToken);
                            state.Step6Completed = true;
                            // Step7: 自動Excel出力（可能なら）
                            if (!state.Step7Completed)
                            {
                                await TryStep7ExcelAsync(turnContext, state, cancellationToken);
                            }
                        }
                    }

                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    return;
                }

                // 未確定: 次の制約確認質問
                var nextConsQ = await GenerateNextConstraintsQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, state.FinalCoreValue, state.Step5QuestionCount, cancellationToken);
                if (string.IsNullOrWhiteSpace(nextConsQ))
                {
                    nextConsQ = await GenerateNextConstraintsQuestionAsync(_kernel, transcript, state.Step1SummaryJson, state.FinalTarget, state.FinalUsageContext, state.FinalCoreValue, state.Step5QuestionCount, cancellationToken);
                }
                if (!string.IsNullOrWhiteSpace(nextConsQ))
                {
                    state.History.Add($"Assistant: {nextConsQ}");
                    TrimHistory(state.History, 30);
                    state.Step5QuestionCount++;
                    if (_conversationState != null)
                    {
                        var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                        await accessor.SetAsync(turnContext, state, cancellationToken);
                        await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
                    }
                    Console.WriteLine($"[USER_MESSAGE] {nextConsQ}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(nextConsQ, nextConsQ), cancellationToken);
                }
                return;
            }

            // Step6（クリエイティブ要素自動生成）単独トリガ（万が一 Step5 後に未実行の場合）
            if (state.Step1Completed && state.Step2Completed && state.Step3Completed && state.Step4Completed && state.Step5Completed && !state.Step6Completed)
            {
                var step6 = await GenerateCreativeElementsAsync(
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
                if (!string.IsNullOrWhiteSpace(step6))
                {
                    Console.WriteLine($"[USER_MESSAGE] {step6}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(step6, step6), cancellationToken);
                    var followMsg = "この要素を使ってキャッチフレーズを作ります。";
                    Console.WriteLine($"[USER_MESSAGE] {followMsg}");
                    await turnContext.SendActivityAsync(MessageFactory.Text(followMsg, followMsg), cancellationToken);
                    state.Step6Completed = true;
                    // Step7: 自動Excel出力（可能なら）
                    if (!state.Step7Completed)
                    {
                        await TryStep7ExcelAsync(turnContext, state, cancellationToken);
                    }
                }
                if (_conversationState != null)
                {
                    var accessor = _conversationState.CreateProperty<ElicitationState>(nameof(ElicitationState));
                    await accessor.SetAsync(turnContext, state, cancellationToken);
                    await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
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

                var summaryMsg = BuildConsolidatedSummary(state);
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
    private static async Task<string?> GenerateNextTargetQuestionAsync(Kernel kernel, string transcript, string? step1SummaryJson, int questionCount, CancellationToken ct)
    {
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

    var response = await InvokeAndLogAsync(kernel, history, ct, "S2/TARGET:Q");
        return response?.Content?.Trim();
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
    private static async Task<string?> GenerateNextMediaQuestionAsync(Kernel kernel, string transcript, string? step1SummaryJson, int questionCount, CancellationToken ct)
    {
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
    var response = await InvokeAndLogAsync(kernel, history, ct, "S3/MEDIA:Q");
        return response?.Content?.Trim();
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
                        history.AddSystemMessage("""
        あなたは日本語コピーライティング支援アシスタントです。以下の確定情報を踏まえて
        『要素アイデア』を JSON 形式のみで出力してください。余計な説明やコードフェンス、前後テキストは禁止です。

        JSON スキーマ（厳守）例:
        {
            "状況": [ "例1", "例2", "例3", "例4", "例5" ],
            "課題・欲求": [ "例1", "例2", "例3", "例4", "例5" ],
            "感情": [ "例1", "例2", "例3", "例4", "例5" ],
            "温度感": [ "例1", "例2", "例3", "例4", "例5" ]
        }

        ガイド:
        - 既出情報と矛盾する具体名や数字を作らない
        - 類似や言い換え重複を避け幅を出す
        - 各配列は必ず 5 要素（不足や過剰禁止）
        - 並び順は意味的に自然なら自由

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
        var response = await InvokeAndLogAsync(kernel, history, ct, "S6/ELEMENTS:GEN");
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
                Console.WriteLine($"[STEP6_JSON_SAVED] {path}");
                // 人間向けにレンダリングして返す
                return RenderCreativeElementsForUser(dict);
            }
            catch (Exception exWrite)
            {
                Console.WriteLine($"[STEP6_JSON_WRITE_ERROR] {exWrite.Message}");
                return RenderCreativeElementsForUser(dict); // 保存失敗でも表示は行う
            }
        }
        Console.WriteLine("[STEP6_JSON_TOTAL_FAIL] JSON/旧形式いずれも解析不可 -> 生raw返却");
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
    private static string GenerateInitialConstraintQuestion()
    {
        return "キャッチコピー作成で考慮すべき制約事項はありますか？例えば『最大◯文字』『避けたい表現』『法律上必要な表記』『文化的に配慮したい点』などがあれば教えてください。なければ『特にない』でOKです。";
    }

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
        CancellationToken ct)
    {
        var history = new ChatHistory();
        var pacing = questionCount >= 2 ? "（質問が続いているため簡潔に。YES/NOか選択肢提示で手短に）" : string.Empty;
        history.AddSystemMessage(@$"あなたは丁寧なコピー制作アシスタントです。今は『制約事項』（文字数 / 文化・配慮 / 法規・レギュレーション / その他）だけを確認します。
        既出を繰り返しすぎない。新規に想像で条件を作らない。{pacing}
        出力は次の1質問のみ（日本語）。");
        var ctx = new List<string>();
        if (!string.IsNullOrWhiteSpace(step1SummaryJson)) ctx.Add(step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[Target]{finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[Usage]{finalUsageContext}");
        if (!string.IsNullOrWhiteSpace(finalCoreValue)) ctx.Add($"[Core]{finalCoreValue}");
        history.AddUserMessage($"--- コンテキスト（参考） ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S5/CONSTRAINTS:Q");
        return response?.Content?.Trim();
    }

    private static string RenderConstraintSummary(ElicitationState state)
    {
        var lines = new List<string>();
        lines.Add("— 制約事項まとめ —");
        lines.Add($"文字数: {(!string.IsNullOrWhiteSpace(state.ConstraintCharacterLimit) ? state.ConstraintCharacterLimit : "指定なし")}");
        lines.Add($"文化・配慮: {(!string.IsNullOrWhiteSpace(state.ConstraintCultural) ? state.ConstraintCultural : "指定なし")}");
        lines.Add($"法規・レギュレーション: {(!string.IsNullOrWhiteSpace(state.ConstraintLegal) ? state.ConstraintLegal : "指定なし")}");
        lines.Add($"その他: {(!string.IsNullOrWhiteSpace(state.ConstraintOther) ? state.ConstraintOther : "指定なし")}");
        return string.Join("\n", lines);
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
        var pacing = questionCount >= 2
            ? "（質問が続いているため、深掘りは控えめに。2〜3個の仮説候補を各1行で示し、『どれが近い？／他にありますか？』とだけ確認）"
            : string.Empty;
        var reasonSnippet = string.IsNullOrWhiteSpace(evalReason) ? "" : $"直前の評価で不足とされた点: {evalReason}\n";
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『提供価値・差別化のコア』だけを確かめます。
            {reasonSnippet}
            既出の手がかり（Step1要約のcoreDriver/subjectOrHero/essenceHints、確定済みのターゲットや媒体、会話内の記述）を尊重し、重複確認は手短に。許可ベースで短く1つだけ質問してください{pacing}。

            コア以外（目的/ターゲット/媒体の再確認、表現案、制作条件など）は扱いません（別フェーズ）。

            必要に応じて2〜3個の候補（各1行）を示し、『どれが近いですか？／他にありますか？』と軽く確認するのはOKです。

            出力はユーザーにそのまま見せる日本語のテキストのみ。評価理由をそのまま繰り返し羅列するのは避け、質問に自然に反映してください。");

        var ctx = new List<string>();
        ctx.Add(string.IsNullOrWhiteSpace(step1SummaryJson) ? "(Step1要約なし)" : step1SummaryJson!);
        if (!string.IsNullOrWhiteSpace(finalTarget)) ctx.Add($"[FinalTarget] {finalTarget}");
        if (!string.IsNullOrWhiteSpace(finalUsageContext)) ctx.Add($"[FinalUsageContext] {finalUsageContext}");
        history.AddUserMessage($"--- コンテキスト（参考） ---\n{string.Join("\n", ctx)}\n\n--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S4/CORE:Q");
        return response?.Content?.Trim();
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

    private static async Task<string?> GenerateNextQuestionAsync(Kernel kernel, string transcript, int questionCount, CancellationToken ct)
    {
    var history = new ChatHistory();
    var pacing = questionCount >= 2
        ? "（すでに質問が続いているため、深掘りは控えめに。2〜3個の方向性の仮説を各1行で提案し、『どれが近い？／他にありますか？』とだけ確認。畳みかけはNG）"
        : string.Empty;
        history.AddSystemMessage(@$"あなたはやさしく話すマーケティングのプロです。工程名は出さず、自然な対話で『活動目的』だけを確かめます。
            テンプレ口調は避け、許可ベースで、短く1つだけ質問してください。質問密度は控えめにしてください{pacing}。

            目的の深度要件（不足があればまずそこを1ステップで聞く）:
            - イベント告知 → 何のイベントか（種類/名称/テーマ）。例: 地域◯◯フェス / 学園祭 / 新製品発表会
            - 販促キャンペーン → 何の製品・サービス・プランか
            - ブランド認知 → どのブランド / どのサービス領域か
            - 採用 → そのままでも可（必要なら職種や層を任意確認）

            既にカテゴリは出ていて “何の◯◯か” が欠けている場合は、それを一問で埋める質問を作成。まだカテゴリ自体が曖昧なら、イベント/販促/ブランド認知/採用/その他 のどれが近いか軽い候補提示も可（最大3行）。

            ターゲット像や表現案、制作条件には踏み込みません（別フェーズ）。

            出力はユーザーにそのまま見せる日本語テキストのみ。余計な前置きや工程名は禁止。");
        history.AddUserMessage($"--- 会話履歴 ---\n{transcript}");
    var response = await InvokeAndLogAsync(kernel, history, ct, "S1/PURPOSE:Q");
        return response?.Content?.Trim();
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
        lines.Add("— ここまでの整理 —");
        if (!string.IsNullOrWhiteSpace(purpose)) lines.Add($"目的: {purpose}");
        if (state.Step2Completed && !string.IsNullOrWhiteSpace(target)) lines.Add($"ターゲット: {target}");
        if (state.Step3Completed && !string.IsNullOrWhiteSpace(usage)) lines.Add($"媒体/利用シーン: {usage}");
        if (state.Step4Completed && !string.IsNullOrWhiteSpace(core)) lines.Add($"コア価値: {core}");
        if (state.Step5Completed)
        {
            lines.Add("制約事項:");
            lines.Add($"- 文字数: {(string.IsNullOrWhiteSpace(charLimit) ? "指定なし" : charLimit)}");
            lines.Add($"- 文化・配慮: {(string.IsNullOrWhiteSpace(cultural) ? "指定なし" : cultural)}");
            lines.Add($"- 法規・レギュレーション: {(string.IsNullOrWhiteSpace(legal) ? "指定なし" : legal)}");
            lines.Add($"- その他: {(string.IsNullOrWhiteSpace(other) ? "指定なし" : other)}");
        }
        return string.Join("\n", lines);
    }

    // Core質問の最低限バリデーション（質問になっているか / 不要な了承だけで終わっていないか）
    private static bool IsInvalidCoreQuestion(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return true;
        var t = text.Trim();
        if (!t.Contains('？') && !t.Contains('?')) return true; // 必ず質問記号
        string[] badStarts = { "了解しました", "承知しました", "わかりました", "ありがとうございます", "ではこれらの魅力を中心に" };
        if (badStarts.Any(bs => t.StartsWith(bs, StringComparison.Ordinal))) return true;
        string[] completionLike = { "考えていきます", "進めます", "作っていきます" };
        if (completionLike.Any(c => t.Contains(c)) && (t.EndsWith("。") || t.EndsWith("！") || t.EndsWith("です。"))) return true;
        return false;
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

        // 十分な候補がなければ汎用一問
        if (candidateList.Count < 2)
        {
            return "提供価値・差別化の核を一言で教えてください。（例：初心者でもすぐ参加できる安心感 など）";
        }

        // 先頭4件のみ採用
        var condensed = candidateList.Take(4).ToList();
        // 番号付け
        var numbered = condensed.Select((c,i)=>$"{i+1}) {c}");
        return "価値の核として最も打ち出したいのはどれでしょうか？ " + string.Join(" ", numbered) + " 5) 他（自由記述）  番号または一言で教えてください。";
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
        var dict = new Dictionary<string, List<string>>
        {
            ["状況"] = new List<string>(),
            ["課題・欲求"] = new List<string>(),
            ["感情"] = new List<string>(),
            ["温度感"] = new List<string>()
        };

        string? current = null;
        var lines = raw.Replace("\r", string.Empty).Split('\n');
        var categoryMap = new Dictionary<string, string>
        {
            ["状況"] = "状況",
            ["課題・欲求"] = "課題・欲求",
            ["感情"] = "感情",
            ["温度感"] = "温度感",
            ["トーン＆キャラクターおよび温度感"] = "温度感" // 旧名称互換
        };

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
            string[] keys = { "状況", "課題・欲求", "感情", "温度感" };
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
        var parts = new List<string?>
        {
            Format("状況"),
            Format("課題・欲求"),
            Format("感情"),
            Format("温度感")
        }.Where(s => !string.IsNullOrWhiteSpace(s));
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

    // Step7: クリエイティブ要素生成後に自動 Excel 出力（OneDrive）
    private async Task TryStep7ExcelAsync(ITurnContext turnContext, ElicitationState state, CancellationToken cancellationToken)
    {
        if (_oneDriveExcelService == null)
        {
            var skip = "（Excel出力は未構成のためスキップしました。/excel で手動コマンド、または OneDrive 環境変数を設定してください。）";
            Console.WriteLine($"[USER_MESSAGE] {skip}");
            await turnContext.SendActivityAsync(MessageFactory.Text(skip, skip), cancellationToken);
            return;
        }
        try
        {
            var pre = "OneDriveにExcelを作成しています。";
            Console.WriteLine($"[USER_MESSAGE] {pre}");
            await turnContext.SendActivityAsync(MessageFactory.Text(pre, pre), cancellationToken);

            string? uploadedUrl = null;
            var progress = new Progress<string>(url =>
            {
                uploadedUrl = url;
                var upMsg = $"アップロード完了（これから内容を書き込みます）: {url}";
                Console.WriteLine($"[USER_MESSAGE] {upMsg}");
                // Fire & forget (TurnContext はスレッドセーフでないので同期呼び出し) → Task.Runせず直接待機
                turnContext.SendActivityAsync(MessageFactory.Text(upMsg, upMsg), cancellationToken).GetAwaiter().GetResult();
            });

            var result = await _oneDriveExcelService.CreateAndFillExcelAsync(progress, cancellationToken);
            if (result.IsSuccess && !string.IsNullOrWhiteSpace(result.WebUrl))
            {
                var done = $"Excel出力完了: {result.WebUrl}";
                Console.WriteLine($"[USER_MESSAGE] {done}");
                await turnContext.SendActivityAsync(MessageFactory.Text(done, done), cancellationToken);
                state.Step7Completed = true;
            }
            else
            {
                var err = $"Excel出力失敗: {result.Error ?? "不明なエラー"}";
                Console.WriteLine($"[USER_MESSAGE] {err}");
                await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
            }
        }
        catch (Exception ex)
        {
            var err = $"Excel出力中に例外: {ex.Message}";
            Console.WriteLine($"[USER_MESSAGE] {err}");
            await turnContext.SendActivityAsync(MessageFactory.Text(err, err), cancellationToken);
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
    public bool Step5Completed { get; set; } = false; // 制約事項（文字数/文化/法規/その他）
    public bool Step6Completed { get; set; } = false; // クリエイティブ要素自動生成
    public bool Step7Completed { get; set; } = false; // Excel出力
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