using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using MicrosoftGraphBot.Models;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphBot.Dialog.ResourceTypes
{
    [Serializable]
    public class TasksDialog : IDialog<bool>
    {
        /// <summary>
        /// Called to start a dialog
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <returns></returns>
#pragma warning disable 1998
        public async Task StartAsync(IDialogContext context)
#pragma warning restore 1998
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
            await ProcessAsync(context);
        }

        private static async Task ProcessAsync(IDialogContext context)
        {
            // Get needed data for the HTTP request.
            var entity = context.ConversationData.GetDialogEntity();
            var requestUrl = $"https://graph.microsoft.com/beta/users/{entity.id}/tasks";
            var httpClient = new HttpClient();
            var accessToken = await context.GetAccessToken();

            // Save the current endpoint and retrieve the dialog entity.
            context.SaveNavCurrent(requestUrl);

            // Perform the HTTP request.
            var response = await httpClient.MSGraphGET(accessToken, requestUrl);
            var tasks = ((JArray)response["value"]).ToTasksList();

            // Remove completed tasks.
            tasks = new List<PlanTask>(tasks.Where(t => t.PercentComplete < 100));

            // Not getting OData to work on /tasks, limiting in client instead.
            tasks = new List<PlanTask>(tasks.Take(5));

            // TODO: Aggregate above filtering to single method. Current setup is for clarity.

            // Create tasks operations.
            var operations = new List<QueryOperation>();
            operations.AddRange(tasks.Select(t =>
            {
                // Create the text (trimmed if needed).
                var text = t.text.Length <= 20
                    ? t.text
                    : new string(t.text.Take(20).ToArray()).Trim() + "...";
                return new QueryOperation
                {
                    Text = text,
                    Type = OperationType.Tasks,
                    Endpoint = t.navEndpoint,
                    ContextObject = t
                };
            }));

            // Create new task operation.
            operations.Add(new QueryOperation
            {
                Text = "(Create new)",
                Type = OperationType.Create,
                Endpoint = requestUrl
            });

            // Create other operations.
            operations.Add(new QueryOperation
            {
                Text = $"(Other {entity} queries)",
                Type = OperationType.ShowOperations
            });

            // Create start over operation.
            operations.Add(new QueryOperation
            {
                Text = "(Start over)",
                Type = OperationType.StartOver
            });

            // Add paging for up, next, previous.
            operations.InitializePaging(context, response);

            // Allow user to select the operation.
            PromptDialog.Choice(context, async (choiceContext, choiceResult) =>
            {
                // Get choice result.
                var operation = await choiceResult;
                switch (operation.Type)
                {
                    case OperationType.Tasks:
                        // Save the operation for recursive call.
                        choiceContext.ConversationData.SetValue("TaskOperation", operation);

                        // The user selected a task, go to it in navigation stack.
                        choiceContext.NavPushItem(choiceContext.GetNavCurrent());
                        choiceContext.NavPushLevel(); 

                        // Handle the selection.
                        await ProcessTaskAsync(choiceContext, operation);
                        break;
                    case OperationType.ShowOperations:
                        choiceContext.Done(false); //return to parent WITHOUT start over
                        break;
                    case OperationType.StartOver:
                        choiceContext.Done(true); //return to parent WITH start over
                        break;
                }

            }, operations, "What would you like to see next?");
        }

        private static async Task ProcessTaskAsync(IDialogContext context, QueryOperation operation)
        {
            var entity = context.ConversationData.GetDialogEntity();
            
            // Get the task.
            var task = operation.GetContextObjectAs<PlanTask>();

            // Create task operations.
            var operations = new List<QueryOperation>();

            // Create in progress operation.
            if (task.PercentComplete == 0)
            {
                operations.Add(new QueryOperation
                {
                    Text = "(In progress)",
                    Type = OperationType.InProgress,
                    Endpoint = operation.Endpoint,
                    ContextObject = task
                });
            }

            // Create complete operation.
            if (task.PercentComplete != 100)
            {
                operations.Add(new QueryOperation
                {
                    Text = "(Complete)",
                    Type = OperationType.Complete,
                    Endpoint = operation.Endpoint,
                    ContextObject = task
                });
            }

            // Create delete operation.
            operations.Add(new QueryOperation
            {
                Text = "(Delete)",
                Type = OperationType.Delete,
                Endpoint = operation.Endpoint,
                ContextObject = task
            });

            // Create other operations.
            operations.Add(new QueryOperation
            {
                Text = $"(Other {entity} queries)",
                Type = OperationType.ShowOperations
            });

            // Create start over operation.
            operations.Add(new QueryOperation
            {
                Text = "(Start over)",
                Type = OperationType.StartOver
            });

            // Add paging for up, next, previous.
            operations.InitializePaging(context, new JObject());

            // Create prompt text.
            var promptText = $"Task \"{task.Title}\" is ";

            // Set completion status.
            switch (task.PercentComplete)
            {
                case 0:
                    promptText += "not started";
                    break;
                case 100:
                    promptText += "completed";
                    break;
                default:
                    promptText += "in progress";
                    break;
            }
            promptText += ". What would you like to do next?";

            // Allow user to select the operation.
            PromptDialog.Choice(context, async (choiceContext, choiceResult) =>
            {
                // Get choice result.
                switch ((await choiceResult).Type)
                {
                    case OperationType.InProgress:
                        PromptDialog.Confirm(choiceContext, ChangeTaskToInProgressAsync,
                            "Are you sure that you want to set this task as in progress?");
                        break;
                    case OperationType.Complete:
                        PromptDialog.Confirm(choiceContext, ChangeTaskToCompletedAsync,
                            "Are you sure that you want to set this task as completed?");
                        break;
                    case OperationType.Delete:
                        PromptDialog.Confirm(choiceContext, DeleteTaskAsync,
                            "Are you sure that you want to delete the task?");
                        break;
                    case OperationType.Up:
                        // Navigating up to parent, pop the level and then pop the
                        // last query on the parent.
                        choiceContext.NavPopLevel(); //pop level to parent
                        choiceContext.NavPopItem();
                        await ProcessAsync(choiceContext);
                        break;
                    case OperationType.ShowOperations:
                        choiceContext.Done(false); //return to parent WITHOUT start over
                        break;
                    case OperationType.StartOver:
                        choiceContext.Done(true); //return to parent WITH start over
                        break;
                }
            }, operations, promptText);
        }

        private async static Task ChangeTaskToInProgressAsync(IDialogContext context,
            IAwaitable<bool> result)
        {
            await UpdateTaskProgressAsync(context, result, 50);
        }

        private async static Task ChangeTaskToCompletedAsync(IDialogContext context,
            IAwaitable<bool> result)
        {
            await UpdateTaskProgressAsync(context, result, 100);
        }

        private async static Task UpdateTaskProgressAsync(IDialogContext context,
            IAwaitable<bool> result, int percentComplete)
        {
            var confirm = await result;
            var operation = context.ConversationData.Get<QueryOperation>("TaskOperation");

            // Get the task.
            var task = operation.GetContextObjectAs<PlanTask>();

            if (confirm)
            {
                var httpClient = new HttpClient();
                var accessToken = await context.GetAccessToken();

                // Set Authorization header.
                httpClient.DefaultRequestHeaders.Authorization = new 
                    AuthenticationHeaderValue("Bearer", accessToken);

                // Set ETag header.
                var headers = httpClient.DefaultRequestHeaders;
                headers.IfMatch.Add(new EntityTagHeaderValue(task.ETag.Substring(2,
                    task.ETag.Length - 2), true));

                // Update the task.
                await context.PostAsync("Updating task...");
                var response = await httpClient.PatchAsJsonAsync(
                    $"https://graph.microsoft.com/beta/{operation.Endpoint}", new 
                {
                    PercentComplete = percentComplete
                });
                await context.PostAsync(response.IsSuccessStatusCode ? "Task updated!" : "Update failed!");

                // Navigating up to parent, pop the level and then pop the
                // last query on the parent.
                context.NavPopLevel(); //pop level to parent
                context.NavPopItem();
                await ProcessAsync(context);
            }
            else
            {
                //show options again
                await ProcessTaskAsync(context, operation);
            }
        }

        private async static Task DeleteTaskAsync(IDialogContext context,
            IAwaitable<bool> result)
        {
            var confirm = await result;
            var operation = context.ConversationData.Get<QueryOperation>("TaskOperation");

            // Get the task.
            var task = operation.GetContextObjectAs<PlanTask>();

            if (confirm)
            {
                var httpClient = new HttpClient();
                var accessToken = await context.GetAccessToken();

                // Set ETag header.
                var headers = httpClient.DefaultRequestHeaders;
                headers.IfMatch.Add(new EntityTagHeaderValue(task.ETag.Substring(2,
                    task.ETag.Length - 2), true));

                // Delete the task.
                await context.PostAsync("Deleting task...");
                var response = await httpClient.MSGraphDELETE(accessToken,
                    $"https://graph.microsoft.com/beta/{operation.Endpoint}");
                await context.PostAsync(response ? "Task deleted!" : "Delete failed!");

                // Navigating up to parent, pop the level and then pop the
                // last query on the parent.
                context.NavPopLevel(); //pop level to parent
                context.NavPopItem();
                await ProcessAsync(context);
            }
            else
            {
                //show options again
                await ProcessTaskAsync(context, operation);
            }
        }
    }
}