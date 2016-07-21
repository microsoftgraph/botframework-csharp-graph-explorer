using AuthBot;
using MicrosoftGraphBot;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Threading;

namespace MicrosoftGraphBot.Dialog.ResourceTypes
{
    [Serializable]
    public class ManagerDialog : IDialog<bool>
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
            var manager = await getManager(context);

            //build a list of valid operations the user can take
            List<Models.QueryOperation> operations = new List<Models.QueryOperation>();
            if (manager != null)
            {
                operations.Add(new Models.QueryOperation() { Text = String.Format("{0}'s photo", manager.ToString()), Type = Models.OperationType.Photo });
                operations.Add(new Models.QueryOperation() { Text = String.Format("{0}'s manager", manager.ToString()), Type = Models.OperationType.Manager });
                operations.Add(new Models.QueryOperation() { Text = String.Format("{0}'s direct reports", manager.ToString()), Type = Models.OperationType.DirectReports });
                operations.Add(new Models.QueryOperation() { Text = String.Format("(Other {0} queries)", manager.ToString()), Type = Models.OperationType.ChangeDialogEntity });
            }
            operations.Add(new Models.QueryOperation() { Text = String.Format("(Other {0} queries)", entity.ToString()), Type = Models.OperationType.ShowOperations });
            operations.Add(new Models.QueryOperation() { Text = "(Start over)", Type = Models.OperationType.StartOver });

            //prepare the message
            string msg = String.Format("{0} doesn't have a manager listed. What would you like to do next:", entity.text);
            if (manager != null)
                msg = String.Format("{0}'s manager is {1}. What would you like to do next:", entity.text, manager.ToString());

            //Allow the user to select next path
            PromptDialog.Choice(context, async (IDialogContext choiceContext, IAwaitable<Models.QueryOperation> choiceResult) =>
            {
                var option = await choiceResult;

                switch (option.Type)
                {
                    case Models.OperationType.Manager:
                    case Models.OperationType.DirectReports:
                    case Models.OperationType.ChangeDialogEntity:
                        //change the dialog entity to the manager
                        var newEntity = await getManager(choiceContext);
                        var eType = (newEntity.IsMe(choiceContext)) ? Models.EntityType.Me : Models.EntityType.User; 
                        choiceContext.ConversationData.SaveDialogEntity(new Models.BaseEntity(newEntity, eType));

                        //proceed based on selection
                        if (option.Type == Models.OperationType.DirectReports)
                            //forward to the PhotoDialog
                            await choiceContext.Forward(new PhotoDialog(), async (IDialogContext drContext, IAwaitable<bool> drResult) =>
                            {
                                var startOver = await drResult;
                                drContext.Done(startOver); //return to parent based on child start over value
                            }, true, CancellationToken.None);
                        else if (option.Type == Models.OperationType.Manager)
                            await MessageReceivedAsync(choiceContext, null); //call back into this dialog to go up a level
                        else if (option.Type == Models.OperationType.DirectReports)
                            //forward to the DirectReportsDialog
                            await choiceContext.Forward(new DirectReportsDialog(), async (IDialogContext drContext, IAwaitable<bool> drResult) =>
                            {
                                var startOver = await drResult;
                                drContext.Done(startOver); //return to parent based on child start over value
                            }, true, CancellationToken.None);
                        else if (option.Type == Models.OperationType.ChangeDialogEntity)
                            choiceContext.Done(false); //return to partent WITHOUT start over
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

        /// <summary>
        /// Performs the MS Graph query for a manager
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <returns>User</returns>
        private async Task<Models.User> getManager(IDialogContext context)
        {
            //Get the manager for the DialogEntity
            var entity = context.ConversationData.GetDialogEntity();
            HttpClient client = new HttpClient();
            var token = await context.GetAccessToken();
            var results = await client.MSGraphGET(token, String.Format("https://graph.microsoft.com/v1.0/users/{0}/manager", entity.id));
            return results.ToUser();
        }
    }
}
