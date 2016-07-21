using AuthBot;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using MicrosoftGraphBot.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class GroupLookupDialog : IDialog<Group>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            //don't await the message...we won't use it
            PromptDialog.Text(context, async (IDialogContext textContext, IAwaitable<string> textResult) =>
            {
                //get the result and perform the group lookup
                var searchText = await textResult;
                var token = await textContext.GetAccessToken("https://graph.microsoft.com");
                var matches = await Group.Lookup(token, searchText);

                //check the number of matches and respond accordingly
                if (matches.Count == 0)
                {
                    //no matches...allow retry
                    PromptDialog.Confirm(textContext, async (IDialogContext retryContext, IAwaitable<bool> retryResult) =>
                    {
                        //check retry result and handle accordingly
                        var retry = await retryResult;
                        if (retry)
                            await MessageReceivedAsync(retryContext, null); //retry
                        else
                            retryContext.Done<User>(null); //return null
                    }, "No matches found...want to try again?");
                }
                else if (matches.Count == 1)
                {
                    //resolve the exact match
                    textContext.Done<Group>(matches[0]);
                }
                else
                {
                    //multiple matches...give choice
                    PromptDialog.Choice(textContext, async (IDialogContext choiceContext, IAwaitable<Group> choiceResult) =>
                    {
                        var selection = await choiceResult;
                        choiceContext.Done(selection);
                    }, matches, "Multiple matches found...which user would you like to explore?");
                }
            }, "What group are you interested in (you can lookup by full/partial name or alias)?");
        }
    }
}
