using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
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
        private const int PageSize = 5;

        public TasksDialog()
        {
            
        }

        /// <summary>
        /// Called to start a dialog.
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <returns></returns>
#pragma warning disable 1998
        public async Task StartAsync(IDialogContext context)
#pragma warning restore 1998
        {
            context.Wait(ShowOperationsAsync);
        }

        /// <summary>
        /// Processes messages received on new thread.
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="item">Awaitable IMessageActivity</param>
        /// <returns>Task</returns>
        public async Task ShowOperationsAsync(IDialogContext context,
            IAwaitable<IMessageActivity> item = null)
        {
            await ShowOperationsAsync(context, 0);
        }

        /// <summary>
        /// Processes messages received on new thread.
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="page">int</param>
        /// <returns>Task</returns>
        public async Task ShowOperationsAsync(IDialogContext context, int page)
        {
            // OData does currently not work well on /tasks (bug). Implementing a
            // custom navigation model.

            // Save the current page.
            context.ConversationData.SetValue("Page", page);

            // Get needed data for the HTTP request.
            var entity = context.ConversationData.GetDialogEntity();
            var requestUrl = (entity.entityType == EntityType.Me || entity.entityType == EntityType.User)
                ? $"https://graph.microsoft.com/beta/users/{entity.id}/tasks"
                : $"https://graph.microsoft.com/beta/plans/{entity.id}/tasks";
            var httpClient = new HttpClient();
            var accessToken = await context.GetAccessToken();

            // Perform the HTTP request.
            var response = await httpClient.MSGraphGET(accessToken, requestUrl);
            var allTasks = ((JArray)response["value"]).ToTasksList();

            // Remove completed tasks.
            allTasks = new List<PlanTask>(allTasks.Where(t => t.PercentComplete < 100));

            // Not getting OData to work on /tasks, limiting in client instead.
            var tasks = new List<PlanTask>(allTasks
                .OrderBy(t => t.CreatedDateTime)
                .Skip(page * PageSize)
                .Take(PageSize));

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
                Text = "(Create task)",
                Type = OperationType.Create,
                Endpoint = requestUrl
            });

            // Create previous page operation.
            if (page > 0)
            {
                operations.Add(new QueryOperation
                {
                    Text = "(Previous page)",
                    Type = OperationType.Previous,
                    Endpoint = requestUrl
                });
            }

            // Create next page operation.
            if ((page + 1) * PageSize < allTasks.Count)
            {
                operations.Add(new QueryOperation
                {
                    Text = "(Next page)",
                    Type = OperationType.Next,
                    Endpoint = requestUrl
                });
            }

            // Create other operations.
            var me = context.ConversationData.Me();
            operations.Add(new QueryOperation
            {
                Text = $"(Other {me} queries)",
                Type = OperationType.ShowOperations
            });

            // Create start over operation.
            operations.Add(new QueryOperation
            {
                Text = "(Start over)",
                Type = OperationType.StartOver
            });

            // Allow user to select the operation.
            PromptDialog.Choice(context, OnOperationsChoiceDialogResume, operations, 
                "What would you like to see next?");
        }

        private async Task OnOperationsChoiceDialogResume(IDialogContext context,
            IAwaitable<QueryOperation> result)
        {
            var page = context.ConversationData.Get<int>("Page");

            // Get choice result.
            var operation = await result;
            switch (operation.Type)
            {
                case OperationType.Tasks:
                    // Save the operation for recursive call.
                    context.ConversationData.SetValue("TaskOperation", operation);

                    // Handle the selection.
                    await ShowTaskOperationsAsync(context, operation);
                    break;
                case OperationType.Create:
                    // Get the dialog entity and see if we can
                    // skip a step (in case of plan).
                    var dialogEntity = context.ConversationData.GetDialogEntity();
                    if (dialogEntity.entityType == EntityType.Me || dialogEntity.entityType == EntityType.User)
                    {
                        await context.Forward(new PlanLookupDialog(), OnPlanLookupDialogResume, 
                            new Plan(), CancellationToken.None);
                    }
                    else
                    {
                        // Save the plan.
                        context.ConversationData.SetValue("Plan", dialogEntity);

                        // Get a bucket.
                        await context.Forward(new BucketLookupDialog(), OnBucketLookupDialogResume, 
                            new Bucket(), CancellationToken.None);
                    }
                    break;
                case OperationType.Next:
                    // Move to the new page.
                    await ShowOperationsAsync(context, page + 1);
                    break;
                case OperationType.Previous:
                    // Move to the previous page.
                    await ShowOperationsAsync(context, page - 1);
                    break;
                case OperationType.ShowOperations:
                    // Reset the dialog entity.
                    context.ConversationData.SaveDialogEntity(new BaseEntity(
                        context.ConversationData.Me(), EntityType.Me));
                    context.Done(false);
                    break;
                case OperationType.StartOver:
                    // Reset the dialog entity.
                    context.ConversationData.SaveDialogEntity(new BaseEntity(
                        context.ConversationData.Me(), EntityType.Me));
                    context.Done(true);
                    break;
            }
        }

        #region Create Task Operation

        private async Task OnPlanLookupDialogResume(IDialogContext context,
            IAwaitable<Plan> result)
        {
            // Save the plan.
            var plan = await result;
            context.ConversationData.SetValue("Plan", plan);

            // Get a bucket.
            await context.Forward(new BucketLookupDialog(), OnBucketLookupDialogResume, new Bucket(),
                CancellationToken.None);
        }

        private async Task OnBucketLookupDialogResume(IDialogContext context,
            IAwaitable<Bucket> result)
        {
            // Save the bucket.
            var bucket = await result;
            context.ConversationData.SetValue("Bucket", bucket);

            // Get the task.
            PromptDialog.Text(context, OnCreateTaskDialogResume,
                "What is the task that you would like to create?");
        }

        private async Task OnCreateTaskDialogResume(IDialogContext context,
            IAwaitable<string> result)
        {
            // Get data needed to create a new task.
            var text = await result;
            var user = context.ConversationData.Get<User>("Me");
            var plan = context.ConversationData.Get<Plan>("Plan");
            var bucket = context.ConversationData.Get<Bucket>("Bucket");

            // Create the task data.
            var task = new PlanTask
            {
                AssignedTo = user.id,
                PlanId = plan.Id,
                BucketId = bucket.Id,
                Title = text
            };

            var httpClient = new HttpClient();
            var accessToken = await context.GetAccessToken();

            // Create the task.
            await context.PostAsync("Creating task...");
            var response = await httpClient.MSGraphPOST(accessToken,
                "https://graph.microsoft.com/beta/tasks", task);
            await context.PostAsync(response ? "Task created!" : "Creation failed!");

            // Clear data.
            context.ConversationData.RemoveValue("Plan");
            context.ConversationData.RemoveValue("Bucket");

            // Show operations.
            await ShowOperationsAsync(context);
        }

        #endregion

        #region Task Operations 

#pragma warning disable 1998
        private async Task ShowTaskOperationsAsync(IDialogContext context, QueryOperation operation)
#pragma warning restore 1998
        {
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

            // Create up operation.
            operations.Add(new QueryOperation
            {
                Text = "(Up)",
                Type = OperationType.Up,
                Endpoint = operation.Endpoint
            });

            // Create other operations.
            var me = context.ConversationData.Me();
            operations.Add(new QueryOperation
            {
                Text = $"(Other {me} queries)",
                Type = OperationType.ShowOperations
            });

            // Create start over operation.
            operations.Add(new QueryOperation
            {
                Text = "(Start over)",
                Type = OperationType.StartOver
            });

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
            PromptDialog.Choice(context, OnTaskOperationsChoiceDialogResume, operations, promptText);
        }

        private async Task OnTaskOperationsChoiceDialogResume(IDialogContext context,
            IAwaitable<QueryOperation> result)
        {
            // Get choice result.
            switch ((await result).Type)
            {
                case OperationType.InProgress:
                    PromptDialog.Confirm(context, OnTaskInProgressDialogResumeAsync,
                        "Are you sure that you want to flag this task as in progress?");
                    break;
                case OperationType.Complete:
                    PromptDialog.Confirm(context, OnTaskCompleteDialogResumeAsync,
                        "Are you sure that you want to flag this task as completed?");
                    break;
                case OperationType.Delete:
                    PromptDialog.Confirm(context, OnDeleteTaskDialogResume,
                        "Are you sure that you want to delete the task?");
                    break;
                    case OperationType.Up:
                    var page = context.ConversationData.Get<int>("Page");
                    await ShowOperationsAsync(context, page);
                    break;
                case OperationType.ShowOperations:
                    // Reset the dialog entity.
                    context.ConversationData.SaveDialogEntity(new BaseEntity(
                        context.ConversationData.Me(), EntityType.Me));
                    context.Done(false);
                    break;
                case OperationType.StartOver:
                    // Reset the dialog entity.
                    context.ConversationData.SaveDialogEntity(new BaseEntity(
                        context.ConversationData.Me(), EntityType.Me));
                    context.Done(true);
                    break;
            }
        }

        #endregion

        #region Change Task Progress Operations

        private async Task OnTaskInProgressDialogResumeAsync(IDialogContext context,
            IAwaitable<bool> result)
        {
            await UpdateTaskProgressAsync(context, result, 50);
        }

        private async Task OnTaskCompleteDialogResumeAsync(IDialogContext context,
            IAwaitable<bool> result)
        {
            await UpdateTaskProgressAsync(context, result, 100);
        }

        private async Task UpdateTaskProgressAsync(IDialogContext context,
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

                // Update the task.
                await context.PostAsync("Updating task...");
                var response = await httpClient.MSGraphPATCH(accessToken,
                    $"https://graph.microsoft.com/beta/{operation.Endpoint}", new
                    {
                        PercentComplete = percentComplete
                    }, task.ETag);
                await context.PostAsync(response ? "Task updated!" : "Update failed!");

                // Show operations.
                await ShowOperationsAsync(context);
            }
            else
            {
                // Show task operations.
                await ShowTaskOperationsAsync(context, operation);
            }
        }

        #endregion

        #region Task Delete Operation

        private async Task OnDeleteTaskDialogResume(IDialogContext context,
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

                // Delete the task.
                await context.PostAsync("Deleting task...");
                var response = await httpClient.MSGraphDELETE(accessToken,
                    $"https://graph.microsoft.com/beta/{operation.Endpoint}", task.ETag);
                await context.PostAsync(response ? "Task deleted!" : "Delete failed!");

                // Navigating up to parent, pop the level and then pop the
                // last query on the parent.
                context.NavPopLevel(); 
                context.NavPopItem();
                await ShowOperationsAsync(context);
            }
            else
            {
                // Show task operations.
                await ShowTaskOperationsAsync(context, operation);
            }
        }

        #endregion
    }
}