using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Dialog.ResourceTypes
{
    [Serializable]
    public class FilesDialog : IDialog<bool>
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
            await processFiles(context, String.Format("https://graph.microsoft.com/v1.0/users/{0}/drive/root/children?$top=5", entity.id));
        }

        /// <summary>
        /// Processes a OneDrive container (root or folder)
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="endpoint">endpoint to query</param>
        /// <returns>Task</returns>
        private async Task processFiles(IDialogContext context, string endpoint)
        {
            //save the current endpoint and retrieve the Dialog Entity
            context.SaveNavCurrent(endpoint);
            var entity = context.ConversationData.GetDialogEntity();

            //perform the http request
            HttpClient client = new HttpClient();
            var token = await context.GetAccessToken();
            JObject json = await client.MSGraphGET(token, endpoint);
            var files = ((JArray)json["value"]).ToFileList();

            //build a list of valid operations the user can take
            List<Models.QueryOperation> operations = new List<Models.QueryOperation>();
            foreach (var file in files)
            {
                //check if file or folder
                if (file.itemType == Models.ItemType.Folder)
                {
                    var folderEndpoint = String.Format("https://graph.microsoft.com/v1.0/users/{0}/drive/items/{1}/children?$top=5", entity.id, file.id);
                    operations.Add(new Models.QueryOperation() { Text = file.ToString(), Type = Models.OperationType.Folder, Endpoint = folderEndpoint });
                }
                else
                    operations.Add(new Models.QueryOperation() { Text = file.ToString(), Type = Models.OperationType.Files, Endpoint = file.navEndpoint });
            }

            //allow users to upload into their own OneDrive
            if (entity.entityType == Models.EntityType.Me)
                operations.Add(new Models.QueryOperation() { Text = "(Upload)", Type = Models.OperationType.Upload, Endpoint = endpoint.Substring(0, endpoint.IndexOf("/children")) });

            //add paging for up, next, prev
            operations.InitializePaging(context, json);

            //add other operations and start over
            operations.Add(new Models.QueryOperation() { Text = String.Format("(Other {0} queries)", entity.ToString()), Type = Models.OperationType.ShowOperations });
            operations.Add(new Models.QueryOperation() { Text = "(Start over)", Type = Models.OperationType.StartOver });

            //Allow the user to select next path
            PromptDialog.Choice(context, async (IDialogContext choiceContext, IAwaitable<Models.QueryOperation> choiceResult) =>
            {
                //CONTEXT SWITCH TO choiceContext
                var operation = await choiceResult;

                switch (operation.Type)
                {
                    case Models.OperationType.Files:
                        //save the operation for recursive call
                        choiceContext.ConversationData.SetValue<Models.QueryOperation>("FileOperation", operation);

                        //the user selected a file...go to it in navstack
                        choiceContext.NavPushItem(choiceContext.GetNavCurrent());
                        choiceContext.NavPushLevel(); //push level to children

                        //handle the selection
                        await processFileSelection(choiceContext, operation);
                        break;
                    case Models.OperationType.Folder:
                        choiceContext.NavPushItem(choiceContext.GetNavCurrent());
                        choiceContext.NavPushLevel(); //push level to children
                        await processFiles(choiceContext, operation.Endpoint);
                        break;
                    case Models.OperationType.Upload:
                        //set upload context
                        choiceContext.ConversationData.SetValue<string>("UploadContext", operation.Endpoint);
                        PromptDialog.Attachment(choiceContext, async (IDialogContext attachmentContext, IAwaitable<IEnumerable<Attachment>> attachmentResult) =>
                        {
                            //CONTEXT SWITCH TO attachmentsContext
                            var attachments = await attachmentResult;

                            //process attachements
                            var uploadToken = await attachmentContext.GetAccessToken();
                            var uploadEndpoint = attachmentContext.ConversationData.Get<string>("UploadContext");
                            foreach (var attachment in attachments)
                            {
                                //parse the filename
                                var filename = attachment.ContentUrl.Substring(attachment.ContentUrl.LastIndexOf("%5c") + 3);

                                //perform the upload and give status along the way
                                await attachmentContext.PostAsync(String.Format("Uploading {0}...", filename));
                                var success = await upload(uploadToken, String.Format("{0}:/{1}:/content", uploadEndpoint, filename), attachment.ContentUrl);
                                if (success)
                                    await attachmentContext.PostAsync(String.Format("{0} uploaded!", filename));
                                else
                                    await attachmentContext.PostAsync(String.Format("{0} upload failed!", filename));
                            }

                            //re-run the current query
                            await processFiles(attachmentContext, attachmentContext.GetNavCurrent());
                        }, "Please select file(s) to upload.");
                        break;
                    case Models.OperationType.Up:
                        //navigating up to parent...pop the level and then pop the last query on the parent
                        choiceContext.NavPopLevel(); //pop level to parent
                        await processFiles(choiceContext, choiceContext.NavPopItem());
                        break;
                    case Models.OperationType.Next:
                        choiceContext.NavPushItem(choiceContext.GetNavCurrent());
                        await processFiles(choiceContext, operation.Endpoint);
                        break;
                    case Models.OperationType.Previous:
                        choiceContext.NavPopItem();
                        await processFiles(choiceContext, operation.Endpoint);
                        break;
                    case Models.OperationType.ShowOperations:
                        choiceContext.Done(false); //return to parent WITHOUT start over
                        break;
                    case Models.OperationType.StartOver:
                        choiceContext.Done(true); //return to parent WITH start over
                        break;
                }
            }, operations, "What would you like to see next?");
        }

        /// <summary>
        /// processes the selection of a file in a OneDrive container
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="operation">QueryOperation</param>
        /// <returns>Task</returns>
        private async Task processFileSelection(IDialogContext context, Models.QueryOperation operation)
        {
            var token = await context.GetAccessToken();
            var entity = context.ConversationData.GetDialogEntity();
            HttpClient client = new HttpClient();
            var file = (await client.MSGraphGET(token, String.Format("https://graph.microsoft.com/v1.0/users/{0}{1}", entity.id, operation.Endpoint))).ToFile();

            //display the file and show new options
            var fileOperations = new List<Models.QueryOperation>();
            fileOperations.Add(new Models.QueryOperation() { Text = "(Delete)", Type = Models.OperationType.Delete, Endpoint = operation.Endpoint });
            fileOperations.Add(new Models.QueryOperation() { Text = "(Download)", Type = Models.OperationType.Download, Endpoint = file.webUrl });

            //add paging for up (empty JObject will leave off next/prev
            fileOperations.InitializePaging(context, new JObject());

            //add other operations and start over
            fileOperations.Add(new Models.QueryOperation() { Text = String.Format("(Other {0} queries)", entity.ToString()), Type = Models.OperationType.ShowOperations });
            fileOperations.Add(new Models.QueryOperation() { Text = "(Start over)", Type = Models.OperationType.StartOver });

            //let the user choose what is next
            PromptDialog.Choice(context, async (IDialogContext fileContext, IAwaitable<Models.QueryOperation> fileResult) =>
            {
                //CONTEXT SWITCH TO fileContext
                var subOperation = await fileResult;
                switch (subOperation.Type)
                {
                    case Models.OperationType.Delete:
                        PromptDialog.Confirm(fileContext, async (IDialogContext confirmContext, IAwaitable<bool> confirmResult) =>
                        {
                            //CONTEXT SWITCH TO confirmContext
                            var confirm = await confirmResult;
                            var fileOp = confirmContext.ConversationData.Get<Models.QueryOperation>("FileOperation");
                            if (confirm)
                            {
                                //delete the file
                                HttpClient deleteClient = new HttpClient();
                                var deleteToken = await confirmContext.GetAccessToken();
                                var deleteEntity = confirmContext.ConversationData.GetDialogEntity();
                                await confirmContext.PostAsync("Deleting file...");
                                var deleteSuccess = await deleteClient.MSGraphDELETE(deleteToken, String.Format("https://graph.microsoft.com/v1.0/users/{0}{1}", deleteEntity.id, fileOp.Endpoint));
                                if (deleteSuccess)
                                    await confirmContext.PostAsync("File deleted!");
                                else
                                    await confirmContext.PostAsync("Delete failed!");

                                //navigating up to parent...pop the level and then pop the last query on the parent
                                confirmContext.NavPopLevel(); //pop level to parent
                                await processFiles(confirmContext, confirmContext.NavPopItem());
                            }
                            else
                            {
                                //show options again
                                await processFileSelection(confirmContext, fileOp);
                            }
                        }, "Are you sure you want to delete the file?");
                        break;
                    case Models.OperationType.Download:
                        //display download link and then call show file options again
                        await fileContext.PostAsync(String.Format("[{0}]({0})", subOperation.Endpoint));
                        await processFileSelection(fileContext, fileContext.ConversationData.Get<Models.QueryOperation>("FileOperation"));
                        break;
                    case Models.OperationType.Up:
                        //navigating up to parent...pop the level and then pop the last query on the parent
                        fileContext.NavPopLevel(); //pop level to parent
                        await processFiles(fileContext, fileContext.NavPopItem());
                        break;
                    case Models.OperationType.ShowOperations:
                        fileContext.Done(false); //return to parent WITHOUT start over
                        break;
                    case Models.OperationType.StartOver:
                        fileContext.Done(true); //return to parent WITH start over
                        break;
                }
            }, fileOperations, String.Format("{0} is {1} bytes. What would you like to see next?", file.text, file.size)); //TODO: show thumbnail???
        }

        /// <summary>
        /// Performs an upload of a file into a OneDrive container
        /// </summary>
        /// <param name="token">AccessToken string</param>
        /// <param name="endpoint">endpoint of container to upload into</param>
        /// <param name="path">path of the local cached file from the upload</param>
        /// <returns>bool for success</returns>
        private async Task<bool> upload(string token, string endpoint, string path)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata.metadata=full");

            //read the file into a stream
            WebRequest readReq = WebRequest.Create(path);
            WebResponse readRes = readReq.GetResponse();
            using (var stream = readRes.GetResponseStream())
            {
                //prepare the content body
                var fileContent = new StreamContent(stream);
                fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                var uploadRes = await client.PutAsync(endpoint, fileContent);
                return uploadRes.IsSuccessStatusCode;
            }
        }
    }
}
