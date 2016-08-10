using AuthBot;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using MicrosoftGraphBot.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MicrosoftGraphBot.Dialog.ResourceTypes;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class GraphDialog : IDialog<string>
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
            //start by having the user select the entity type to query
            await context.Initialize();
            PromptDialog.Choice(context, this.EntityTypeSelected, Enum.GetNames(typeof(EntityType)), "Where do you want to start exploring?");
        }

        /// <summary>
        /// Resume from user selecting the type of entity they want to query
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="item">Awaitable string</param>
        /// <returns>Task</returns>
        public async Task EntityTypeSelected(IDialogContext context, IAwaitable<string> item)
        {
            EntityType intent = (EntityType)Enum.Parse(typeof(EntityType), await item);
            switch (intent)
            {
                case EntityType.Me:
                    //get me
                    var me = context.ConversationData.Me();

                    //save the entity and show available operations
                    context.ConversationData.SaveDialogEntity(new BaseEntity(me, EntityType.Me));
                    await routeOperation(context);
                    break;
                case EntityType.User:
                    //prompt for user
                    await context.Forward(new UserLookupDialog(), async (IDialogContext lookupContext, IAwaitable<User> lookupResult) =>
                    {
                        var user = await lookupResult;

                        //save the entity and show available operations
                        lookupContext.ConversationData.SaveDialogEntity(new BaseEntity(user, EntityType.User));
                        await routeOperation(lookupContext);
                    }, new User(), CancellationToken.None);
                    break;
                case EntityType.Group:
                    //prompt for group
                    await context.Forward(new GroupLookupDialog(), async (IDialogContext lookupContext, IAwaitable<Group> lookupResult) =>
                    {
                        var group = await lookupResult;

                        //save the entity and show available operations
                        lookupContext.ConversationData.SaveDialogEntity(new BaseEntity(group));
                        await routeOperation(lookupContext);
                    }, new Group(), CancellationToken.None);
                    break;
            }
        }

        /// <summary>
        /// Prompts the user to chose an operation and routes it to the appropriate dialog
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <returns>Task</returns>
        private async Task routeOperation(IDialogContext context)
        {
            //initialize a new operation on the context
            context.NewOperation();

            //check the entity type to determine the valid operations
            var entity = context.ConversationData.GetDialogEntity();
            List<QueryOperation> operations = QueryOperation.GetEntityResourceTypes(entity.entityType);

            //prepare the prompt
            string prompt = "What would like to lookup for you?";
            if (entity.entityType != EntityType.Me)
                prompt = String.Format("What would like to lookup for {0}?", entity.text);

            //add start over
            operations.Add(new Models.QueryOperation() { Text = "(Start over)", Type = Models.OperationType.StartOver });

            //let the user select an operation
            PromptDialog.Choice(context, async (IDialogContext opContext, IAwaitable<QueryOperation> opResult) =>
            {
                //check the operation selected and route appropriately
                var operation = await opResult;
                switch (operation.Type)
                {
                    case OperationType.StartOver:
                        await this.MessageReceivedAsync(opContext, null);
                        break;
                    case OperationType.Manager:
                        await opContext.Forward(new ResourceTypes.ManagerDialog(), OperationComplete, true, CancellationToken.None);
                        break;
                    case OperationType.DirectReports:
                        await opContext.Forward(new ResourceTypes.DirectReportsDialog(), OperationComplete, true, CancellationToken.None);
                        break;
                    case OperationType.Files:
                        await opContext.Forward(new ResourceTypes.FilesDialog(), OperationComplete, true, CancellationToken.None);
                        break;
                    case OperationType.Members:
                        await opContext.Forward(new ResourceTypes.MembersDialog(), OperationComplete, true, CancellationToken.None);
                        break;
                    case OperationType.Contacts:
                    case OperationType.Conversations:
                    case OperationType.Events:
                    case OperationType.Groups:
                    case OperationType.Mail:
                    case OperationType.Notebooks:
                    case OperationType.People:
                    case OperationType.Photo:
                    case OperationType.Plans:
                        await opContext.Forward(new PlanLookupDialog(), OnPlanLookupDialogResumeAsync, new Plan(), CancellationToken.None);
                        break;
                    case OperationType.Tasks:
                        await opContext.Forward(new TasksDialog(), OperationComplete, true, CancellationToken.None);
                        break;
                    case OperationType.TrendingAround:
                    case OperationType.WorkingWith:
                        await opContext.PostAsync("Operation not yet implemented");
                        opContext.Wait(MessageReceivedAsync);
                        break;
                }

            }, operations, prompt);
        }

        private async Task OnPlanLookupDialogResumeAsync(IDialogContext context, IAwaitable<Plan> result)
        {
            var plan = await result;

        }

        /// <summary>
        /// The resume from performing a Graph operation
        /// Allows the user to start over or select a different operation
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="result">Awaitable bool indicating if start over</param>
        /// <returns>Task</returns>
        public async Task OperationComplete(IDialogContext context, IAwaitable<bool> result)
        {
            var startOver = await result;
            if (startOver)
                await this.MessageReceivedAsync(context, null);
            else
                await this.routeOperation(context);
        }
    }
}
