using AuthBot;
using AuthBot.Dialogs;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class AuthDialog : IDialog<string>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            //original message if we need to resume
            var message = await item;

            //try to get access token
            var token = await context.GetAccessToken("https://graph.microsoft.com/");
            if (string.IsNullOrEmpty(token))
            {
                //invoke the AuthBot to help get authenticated
                await context.Forward(new AzureAuthDialog("https://graph.microsoft.com/"), this.resumeAfterAuth, message, CancellationToken.None);
            }
            else
            {
                //token exists...forward to GraphDialog
                await context.Forward(new GraphDialog(), null, message, CancellationToken.None);
            }
        }

        private async Task resumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
            //post the response message and then go back into the MessageReceivedAsync flow
            var message = await result;
            await context.PostAsync(message);

            //now that token exists...forward to GraphDialog
            await context.Forward(new GraphDialog(), null, message, CancellationToken.None);
        }
    }
}
