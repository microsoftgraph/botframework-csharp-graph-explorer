using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Dialog.ResourceTypes
{
    [Serializable]
    public class MembersDialog : IDialog<bool>
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
            await processMembers(context, String.Format("https://graph.microsoft.com/v1.0/groups/{0}/members?$top=5", entity.id));
 
        }

        private async Task processMembers(IDialogContext context, string endpoint)
        {
            //save the current endpoint and retrieve the Dialog Entity
            context.SaveNavCurrent(endpoint);
            var entity = context.ConversationData.GetDialogEntity();
            var token = await context.GetAccessToken();

            //get the members from the database based on provided endpoint
            HttpClient client = new HttpClient();
            var json = await client.MSGraphGET(token, endpoint);
            var members = ((JArray)json["value"]).ToUserList();

            //convert to operations
            List<Models.QueryOperation> operations = new List<Models.QueryOperation>();
            foreach (var m in members)
                operations.Add(new Models.QueryOperation() { Text = m.text, Type = Models.OperationType.People, Endpoint = String.Format("https://graph.microsoft.com/v1.0/users/{0}", m.id) });
            
            //add paging
            operations.InitializePaging(context, json);

            //add other operations and start over
            operations.Add(new Models.QueryOperation() { Text = String.Format("(Other {0} queries)", entity.ToString()), Type = Models.OperationType.ShowOperations });
            operations.Add(new Models.QueryOperation() { Text = "(Start over)", Type = Models.OperationType.StartOver });

            //prompt the next selection
            PromptDialog.Choice(context, async (IDialogContext choiceContext, IAwaitable<Models.QueryOperation> choiceResult) =>
            {
                var operation = await choiceResult;
                switch (operation.Type)
                {
                    case Models.OperationType.People:
                        //change the dialog entity and return without start over
                        HttpClient c = new HttpClient();
                        var t = await choiceContext.GetAccessToken();
                        var j = await c.MSGraphGET(t, operation.Endpoint);
                        var newEntity = j.ToUser();
                        var etype = (newEntity.IsMe(choiceContext)) ? Models.EntityType.Me : Models.EntityType.User;
                        choiceContext.ConversationData.SaveDialogEntity(new Models.BaseEntity(newEntity, etype));
                        choiceContext.Done(false); //return with false to stick with same user
                        break;
                    case Models.OperationType.Next:
                        choiceContext.NavPushItem(choiceContext.GetNavCurrent());
                        await processMembers(choiceContext, operation.Endpoint);
                        break;
                    case Models.OperationType.Previous:
                        choiceContext.NavPopItem();
                        await processMembers(choiceContext, operation.Endpoint);
                        break;
                    case Models.OperationType.ShowOperations:
                        choiceContext.Done(false); //return with false to stick with same user
                        break;
                    case Models.OperationType.StartOver:
                        choiceContext.Done(true); //return with true to start over
                        break;
                }
            }, operations, String.Format("{0} has the following members:", entity.text));
        }
    }
}
