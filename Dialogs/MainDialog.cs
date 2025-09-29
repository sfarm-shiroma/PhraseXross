using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using PhraseXross.Services;

namespace PhraseXross.Dialogs;

/// <summary>
/// メイン Dialog: Teams SSO (OAuthPrompt) を使ってユーザー委譲トークンを取得し、
/// 取得後は ElicitationState に保存して終了します。Excel 出力自体は SimpleBot 側で継続処理。
/// </summary>
public class MainDialog : ComponentDialog
{
    public const string DialogIdConst = nameof(MainDialog);
    private readonly UserState _userState;

    public MainDialog(UserState userState)
    : base(DialogIdConst)
    {
        _userState = userState;

        var connectionName = Environment.GetEnvironmentVariable("BOT_OAUTH_CONNECTION_NAME") ?? "GraphDelegated";

        // OAuthPrompt 構成
        var oauthPrompt = new OAuthPrompt(
            nameof(OAuthPrompt),
            new OAuthPromptSettings
            {
                ConnectionName = connectionName,
                Text = "OneDrive へのアクセス許可のためサインインしてください。ブラウザが開かない場合は Teams を再読み込みしてください。",
                Title = "サインイン",
                Timeout = 300_000 // 5分
            });

        AddDialog(oauthPrompt);
        AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
        {
            PromptStepAsync,
            AfterPromptAsync
        }));

        InitialDialogId = nameof(WaterfallDialog);
    }

    private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
    {
        Console.WriteLine("[OAUTH][PromptStep] Starting OAuthPrompt (connection check)");
        try
        {
            var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
            var state = await accessor.GetAsync(stepContext.Context, () => ElicitationState.CreateNew(), cancellationToken);
            state.OAuthPromptStartCount++;
            state.OAuthPromptLastAttemptUtc = DateTimeOffset.UtcNow;
            await accessor.SetAsync(stepContext.Context, state, cancellationToken);
            await _userState.SaveChangesAsync(stepContext.Context, false, cancellationToken);
        }
        catch (Exception ex)
        {
            Console.WriteLine("[OAUTH][PromptStep][WARN] State update failed: " + ex.Message);
        }
        return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
    }

    private async Task<DialogTurnResult> AfterPromptAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
    {
        Console.WriteLine("[OAUTH][AfterPrompt] Returned from OAuthPrompt");
        var tokenResponse = stepContext.Result as TokenResponse;
        if (tokenResponse?.Token != null)
        {
            Console.WriteLine("[OAUTH][AfterPrompt] Token acquired length=" + tokenResponse.Token.Length);
            var accessor = _userState.CreateProperty<ElicitationState>(nameof(ElicitationState));
            var state = await accessor.GetAsync(stepContext.Context, () => ElicitationState.CreateNew(), cancellationToken);
            state.DelegatedGraphToken = tokenResponse.Token;
            state.WaitingForSignIn = false;
            state.LastTokenAcquiredUtc = DateTimeOffset.UtcNow;
            await accessor.SetAsync(stepContext.Context, state, cancellationToken);
            await _userState.SaveChangesAsync(stepContext.Context, false, cancellationToken);
            await stepContext.Context.SendActivityAsync(MessageFactory.Text("サインインが完了しました。処理を続行します。"), cancellationToken);
        }
        else
        {
            Console.WriteLine("[OAUTH][AfterPrompt] Token missing (user may have cancelled or card not completed)");
            await stepContext.Context.SendActivityAsync(MessageFactory.Text("サインインが完了しませんでした。必要ならもう一度試してください。"), cancellationToken);
        }
        return await stepContext.EndDialogAsync(null, cancellationToken);
    }
}
