using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Threading;
using System.Threading.Tasks;

public class SimpleBot : ActivityHandler
{
    private readonly Kernel? _kernel; // optional

    public SimpleBot(Kernel? kernel = null)
    {
        _kernel = kernel;
    }

    protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        var text = turnContext.Activity.Text ?? string.Empty;

        if (_kernel != null)
        {
            try
            {
                var chat = _kernel.GetRequiredService<IChatCompletionService>();
                var history = new ChatHistory();
                history.AddSystemMessage("あなたは丁寧で簡潔なアシスタントです。必要に応じて日本語で回答してください。");
                history.AddUserMessage(text);

                var reply = await chat.GetChatMessageContentAsync(history, cancellationToken: cancellationToken);
                var content = reply?.Content ?? string.Empty;
                if (string.IsNullOrWhiteSpace(content))
                {
                    content = $"(SK応答が空でした) You said: {text}";
                }
                await turnContext.SendActivityAsync(MessageFactory.Text(content, content), cancellationToken);
                return;
            }
            catch (System.Exception ex)
            {
                var fb = $"(SKエラー) {ex.Message} | Fallback: You said: {text}";
                await turnContext.SendActivityAsync(MessageFactory.Text(fb, fb), cancellationToken);
                return;
            }
        }

        var replyText = $"You said: {text}";
        await turnContext.SendActivityAsync(MessageFactory.Text(replyText, replyText), cancellationToken);
    }
}