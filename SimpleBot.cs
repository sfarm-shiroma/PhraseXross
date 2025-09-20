using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using System.Threading;
using System.Threading.Tasks;

public class SimpleBot : ActivityHandler
{
    protected override Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
    {
        var replyText = $"You said: {turnContext.Activity.Text}";
        return turnContext.SendActivityAsync(MessageFactory.Text(replyText, replyText), cancellationToken);
    }
}