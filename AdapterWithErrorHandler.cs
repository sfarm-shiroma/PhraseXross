using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;

namespace PhraseXross
{
    public class AdapterWithErrorHandler : CloudAdapter
    {
        public AdapterWithErrorHandler(BotFrameworkAuthentication auth, ILogger<CloudAdapter> logger)
            : base(auth, logger)
        {
            OnTurnError = async (turnContext, exception) =>
            {
                logger.LogError(exception, "[BOT ERROR] Unhandled exception");
                var msg = "エラーが発生しました。後で再度お試しください。";
                await turnContext.SendActivityAsync(msg);
            };
        }
    }
}
