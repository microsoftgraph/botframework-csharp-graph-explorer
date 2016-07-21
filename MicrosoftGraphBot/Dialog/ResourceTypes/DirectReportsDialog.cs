using MicrosoftGraphBot;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphBot.Dialog.ResourceTypes
{
    public class DirectReportsDialog : IDialog<bool>
    {
        /// <summary>
        /// Called to start a dialog
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <returns></returns>
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        /// <summary>
        /// Processes messages received on new thread
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="item">Awaitable IMessageActivity</param>
        /// <returns>Task</returns>
        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var entity = context.ConversationData.GetDialogEntity();
            var directReports = await getDirectReports(context);

            //build a list of valid operations the user can take
            List<Models.QueryOperation> operations = new List<Models.QueryOperation>();
            foreach (var report in directReports)
                operations.Add(new Models.QueryOperation() { Text = report.text, Type = Models.OperationType.ChangeDialogEntity, ContextObject = report });
            operations.Add(new Models.QueryOperation() { Text = String.Format("(Other {0} queries)", entity.ToString()), Type = Models.OperationType.ShowOperations });
            operations.Add(new Models.QueryOperation() { Text = "(Start over)", Type = Models.OperationType.StartOver });

            //prepare the message
            var msg = String.Format("{0} does not have direct reports lists. What would you like to do next:", entity.text);
            if (directReports.Count == 0)
                msg = String.Format("{0} has the following direct reports. What would you like to do next:", entity.text);

            //Allow the user to select next path
            PromptDialog.Choice(context, async (IDialogContext choiceContext, IAwaitable<Models.QueryOperation> choiceResult) =>
            {
                var option = await choiceResult;

                switch (option.Type)
                {
                    case Models.OperationType.ChangeDialogEntity:
                        var user = (Models.User)option.ContextObject;
                        var eType = (user.IsMe(choiceContext)) ? Models.EntityType.Me : Models.EntityType.User;
                        choiceContext.ConversationData.SaveDialogEntity(new Models.BaseEntity(user, eType));
                        choiceContext.Done(false); //return to parent WITHOUT start over
                        break;
                    case Models.OperationType.ShowOperations:
                        choiceContext.Done(false); //return to parent WITHOUT start over
                        break;
                    case Models.OperationType.StartOver:
                        choiceContext.Done(true); //return to parent WITH start over
                        break;
                }
            }, operations, msg);
        }

        private async Task<List<Models.User>> getDirectReports(IDialogContext context)
        {
            //Get the manager for the DialogEntity
            var entity = context.ConversationData.GetDialogEntity();
            HttpClient client = new HttpClient();
            var token = await context.GetAccessToken();
            var results = await client.MSGraphGET(token, String.Format("https://graph.microsoft.com/v1.0/users/{0}/directReports", entity.id));
            return ((JArray)results["value"]).ToUserList();
        }
    }
}
