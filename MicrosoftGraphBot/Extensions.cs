using AuthBot;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
using MicrosoftGraphBot.Dialog;
using MicrosoftGraphBot.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MicrosoftGraphBot
{
    public static class Extensions
    {
        /// <summary>
        /// Simplifies the query for a token given our resource will always be the Microsoft Graph
        /// </summary>
        /// <param name="context">IBotContext</param>
        /// <returns>AccessToken string</returns>
        public static async Task<string> GetAccessToken(this IBotContext context)
        {
            return await context.GetAccessToken("https://graph.microsoft.com");
        }

        /// <summary>
        /// Converts a list of Graph ResourceType enums to a list of QueryOperations
        /// </summary>
        /// <param name="resourceTypes">List of ResourceType enums</param>
        /// <returns>List of QueryOperation objects</returns>
        public static List<QueryOperation> ToQueryOperations(this List<OperationType> resourceTypes)
        {
            List<QueryOperation> operations = new List<QueryOperation>();
            foreach (var op in resourceTypes)
            {
                operations.Add(new QueryOperation()
                {
                    Type = op,
                    Text = Resource.ResourceManager.GetString(op.ToString()),
                    Endpoint = Resource.ResourceManager.GetString(String.Format("{0}_Endpoint", op.ToString()))
                });
            }
            return operations;
        }

        /// <summary>
        /// Saves the dialog entity (who queries are being performed for)
        /// </summary>
        /// <param name="conversationData">IBotDataBag</param>
        /// <param name="entity">BaseEntity</param>
        public static void SaveDialogEntity(this IBotDataBag conversationData, BaseEntity entity)
        {
            conversationData.SetValue<BaseEntity>("DialogEntity", entity);
        }

        /// <summary>
        /// Gets the dialog entity (who queries are being performed for)
        /// </summary>
        /// <param name="conversationData">IBotDataBag</param>
        /// <returns>BaseEntity</returns>
        public static BaseEntity GetDialogEntity(this IBotDataBag conversationData)
        {
            return conversationData.Get<BaseEntity>("DialogEntity");
        }

        /// <summary>
        /// Initializes the ConversationData by clearing it from the last thread
        /// </summary>
        /// <param name="context">IDialogContext</param>
        /// <param name="result">LuisResult</param>
        public static async Task Initialize(this IDialogContext context)
        {
            //reset the conversation data for new thead
            context.ConversationData.RemoveValue("DialogEntity");

            //try to initialize me
            User me = null;
            if (!context.ConversationData.TryGetValue<User>("Me", out me))
            {
                var token = await context.GetAccessToken();
                me = await User.Me(token);
                context.ConversationData.SetValue<User>("Me", me);
            }
        }

        /// <summary>
        /// resets cached data for a new operation
        /// </summary>
        /// <param name="context">IDialogContext</param>
        public static void NewOperation(this IDialogContext context)
        {
            //initialize the NavPagingStack...provides two-dimensional breadcrumb
            //First list is the level...second is the paging
            //ex: level = folder; paging = paging for files in that folder
            //ex: level = person; paging = paged list of direct reports for a person
            var navStack = new List<List<String>>();
            navStack.Add(new List<string>());
            context.ConversationData.SetValue<List<List<string>>>("NavStack", navStack);
            context.ConversationData.RemoveValue("NavCurrent");
        }

        /// <summary>
        /// Gets me
        /// </summary>
        /// <param name="conversationData">IBotDataBag</param>
        /// <returns>User</returns>
        public static User Me(this IBotDataBag conversationData)
        {
            return conversationData.Get<User>("Me");
        }

        /// <summary>
        /// checks if a specific user is ME
        /// </summary>
        /// <param name="user">User to check against</param>
        /// <param name="context">IDialogContext</param>
        /// <returns>bool</returns>
        public static bool IsMe(this User user, IDialogContext context)
        {
            var me = context.ConversationData.Get<User>("Me");
            return me.id.Equals(user.id, StringComparison.CurrentCultureIgnoreCase);
        }

        //START - THESE ARE ALL PAGING UTILITIES
        public static void SaveNavCurrent(this IDialogContext context, string endpoint)
        {
            context.ConversationData.SetValue<string>("NavCurrent", endpoint);
        }

        public static string GetNavCurrent(this IDialogContext context)
        {
            return context.ConversationData.Get<string>("NavCurrent");
        }

        public static void NavPushLevel(this IDialogContext context)
        {
            //implemented as a list and not a real stack because serialization was corrupting order
            var stack = context.ConversationData.Get<List<List<string>>>("NavStack");
            stack.Insert(0, new List<string>());
            context.ConversationData.SetValue<List<List<string>>>("NavStack", stack);
        }

        public static void NavPushItem(this IDialogContext context, string endpoint)
        {
            //implemented as a list and not a real stack because serialization was corrupting order
            var stack = context.ConversationData.Get<List<List<string>>>("NavStack");
            stack[0].Insert(0, endpoint);
            context.ConversationData.SetValue<List<List<string>>>("NavStack", stack);
        }

        public static void NavPopLevel(this IDialogContext context)
        {
            //implemented as a list and not a real stack because serialization was corrupting order
            var stack = context.ConversationData.Get<List<List<string>>>("NavStack");
            stack.RemoveAt(0);
            context.ConversationData.SetValue<List<List<string>>>("NavStack", stack);
        }

        public static string NavPopItem(this IDialogContext context)
        {
            //implemented as a list and not a real stack because serialization was corrupting order
            var stack = context.ConversationData.Get<List<List<string>>>("NavStack");
            var path = stack[0][0];
            stack[0].RemoveAt(0);
            context.ConversationData.SetValue<List<List<string>>>("NavStack", stack);
            return path;
        }

        public static string NavPeekLevel(this IDialogContext context)
        {
            //implemented as a list and not a real stack because serialization was corrupting order
            var stack = context.ConversationData.Get<List<List<string>>>("NavStack");
            if (stack.Count > 1)
                return stack[1][0];
            else
                return null;
        }

        public static string NavPeekItem(this IDialogContext context)
        {
            //implemented as a list and not a real stack because serialization was corrupting order
            var stack = context.ConversationData.Get<List<List<string>>>("NavStack");
            if (stack[0].Count > 0)
                return stack[0][0];
            else
                return null;
        }

        public static void InitializePaging(this List<QueryOperation> operations, IDialogContext context, JObject json)
        {
            //add next link to the end
            if (json["@odata.nextLink"] != null)
            {
                var next = json.Value<string>("@odata.nextLink");
                operations.Add(new QueryOperation() {Text = "(Next page)", Type = OperationType.Next, Endpoint = next});
            }

            //add previous to the front
            if (!String.IsNullOrEmpty(context.NavPeekItem()))
            {
                var prev = context.NavPeekItem();

                operations.Add(new QueryOperation()
                {
                    Text = "(Prev page)",
                    Type = OperationType.Previous,
                    Endpoint = prev
                });
            }

            //add parent nav up
            if (!String.IsNullOrEmpty(context.NavPeekLevel()))
                operations.Add(new QueryOperation()
                {
                    Text = "(Up to parent)",
                    Type = OperationType.Up,
                    Endpoint = context.NavPeekLevel()
                });
        }

        //START - THESE ARE ALL PAGING UTILITIES


        /// <summary>
        /// Performs a simple HTTP GET against the MSGraph given a token and endpoint
        /// </summary>
        /// <param name="client">HttpClient</param>
        /// <param name="token">Access token string</param>
        /// <param name="endpoint">endpoint uri to perform GET on</param>
        /// <returns>JObject</returns>
        public static async Task<JObject> MSGraphGET(this HttpClient client, string token, string endpoint)
        {
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            using (var response = await client.GetAsync(endpoint))
            {
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    return JObject.Parse(json);
                }
                else
                    return null;
            }
        }

        /// <summary>
        /// Performs a HTTP DELETE against the MSGraph given an access token and request URI
        /// </summary>
        /// <param name="httpClient">HttpClient</param>
        /// <param name="accessToken">Access token string</param>
        /// <param name="requestUri">Request URI to perform DELETE on</param>
        /// <param name="weakETag">Entity Tag header for Microsoft Graph item to perform DELETE on</param>
        /// <returns>boolean for success</returns>
        public static async Task<bool> MSGraphDELETE(this HttpClient httpClient, string accessToken,
            string requestUri, string weakETag = null)
        {
            // Set Authorization and Accept header.
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            // Set (weak) If-Match header.
            if (weakETag != null)
            {
                var headers = httpClient.DefaultRequestHeaders;
                headers.IfMatch.Add(new EntityTagHeaderValue(weakETag.Substring(2,
                    weakETag.Length - 2), true));
            }

            using (var response = await httpClient.DeleteAsync(requestUri))
            {
                return response.IsSuccessStatusCode;
            }
        }

        /// <summary>
        /// Performs a HTTP POST against the MSGraph given an access token and request URI
        /// </summary>
        /// <param name="httpClient">HttpClient</param>
        /// <param name="accessToken">Access token string</param>
        /// <param name="requestUri">Request uri to perform POST on</param>
        /// <param name="data">Request body data for the request</param>
        /// <param name="weakETag">Entity Tag header for Microsoft Graph item to perform POST on</param>
        /// <returns>boolean for success</returns>
        public static async Task<bool> MSGraphPOST<T>(this HttpClient httpClient, string accessToken,
            string requestUri, T data, string weakETag = null) where T : class
        {
            // Set Authorization and Accept header.
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            // Set If-Match header.
            if (weakETag != null)
            {
                var headers = httpClient.DefaultRequestHeaders;
                headers.IfMatch.Add(new EntityTagHeaderValue(weakETag.Substring(2,
                    weakETag.Length - 2), true));
            }

            // Create data.
            var json = JsonConvert.SerializeObject(data);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            using (var response = await httpClient.PostAsync(requestUri, content))
            {
                return response.IsSuccessStatusCode;
            }
        }

        /// <summary>
        /// Performs a HTTP PATCH against the MSGraph given an access token and request URI
        /// </summary>
        /// <param name="httpClient">HttpClient</param>
        /// <param name="accessToken">Access token string</param>
        /// <param name="requestUri">Request uri to perform PATCH on</param>
        /// <param name="data">Request body data for the request</param>
        /// <param name="weakETag">Entity Tag header for Microsoft Graph item to perform PATCH on</param>
        /// <returns>boolean for success</returns>
        public static async Task<bool> MSGraphPATCH<T>(this HttpClient httpClient, string accessToken,
            string requestUri, T data, string weakETag = null) where T : class
        {
            // Set Authorization and Accept header.
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            // Set If-Match header.
            if (weakETag != null)
            {
                var headers = httpClient.DefaultRequestHeaders;
                headers.IfMatch.Add(new EntityTagHeaderValue(weakETag.Substring(2,
                    weakETag.Length - 2), true));
            }

            using (var response = await httpClient.PatchAsJsonAsync(requestUri, data))
            {
                return response.IsSuccessStatusCode;
            }
        }

        /// <summary>
        /// Implements the PATCH HTTP Method on the HttpClient.
        /// </summary>
        /// <param name="client">HttpClient</param>
        /// <param name="requestUri">Uri</param>
        /// <param name="value">T</param>
        /// <returns>HttpResponseMessage</returns>
        public static Task<HttpResponseMessage> PatchAsJsonAsync<T>(this HttpClient client,
            string requestUri, T value) where T : class
        {
            var request = new HttpRequestMessage(new HttpMethod("PATCH"), requestUri)
            {
                Content = new StringContent(JsonConvert.SerializeObject(value),
                    Encoding.UTF8, "application/json")
            };
            return client.SendAsync(request);
        }


        /// 
        /// Below are extensions to parse JSON.NET objects to strongly-typed graph objects
        ///

        /// <summary>
        /// Parses JArray to generic List of File objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of File objects</returns>
        public static List<File> ToFileList(this JArray array)
        {
            List<File> files = new List<File>();
            foreach (var item in array)
                files.Add(item.ToFile());
            return files;
        }

        /// <summary>
        /// Parses JToken to File object
        /// </summary>
        /// <param name="array">JToken</param>
        /// <returns>File object</returns>
        public static File ToFile(this JToken obj)
        {
            File f = null;
            if (obj != null)
            {
                f = new File()
                {
                    id = obj.Value<string>("id"),
                    text = obj.Value<string>("name"),
                    size = obj.Value<int>("size"),
                    webUrl = obj.Value<string>("webUrl"),
                    itemType = (obj["folder"] != null) ? ItemType.Folder : ItemType.File,
                    navEndpoint = String.Format("/drive/items/{0}", obj.Value<string>("id"))

                };
            }

            return f;
        }


        /// <summary>
        /// Parses JArray to generic List of File objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of Mail objects</returns>
        public static List<Mail> ToMailList(this JArray array)
        {
            List<Mail> messages = new List<Mail>();
            foreach (var item in array)
                messages.Add(item.ToMail());
            return messages;
        }

        /// <summary>
        /// Parses JToken to Mail object
        /// </summary>
        /// <param name="array">JToken</param>
        /// <returns>Mail object</returns>
        public static Mail ToMail(this JToken obj)
        {
            Mail m = null;
            if (obj != null)
            {
                m = new Mail()
                {
                    id = obj.Value<string>("id"),
                    text = obj.Value<string>("subject"),
                    isRead = obj.Value<bool>("isRead"),
                    senderName = obj.SelectToken("sender.emailAddress").Value<string>("name"),
                    senderEmail = obj.SelectToken("sender.emailAddress").Value<string>("address"),
                    importance = obj.Value<string>("importance"),
                    sentDate = obj.Value<DateTime>("sentDateTime"),
                    itemType = ItemType.Mail
                };
            }

            return m;
        }


        /// <summary>
        /// Parses JArray to generic List of User objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of User objects</returns>
        public static List<User> ToUserList(this JArray array)
        {
            List<User> users = new List<User>();
            foreach (var item in array)
                users.Add(item.ToUser());
            return users;
        }

        /// <summary>
        /// Parses JToken to User object
        /// </summary>
        /// <param name="array">JToken</param>
        /// <returns>User object</returns>
        public static User ToUser(this JToken obj)
        {
            User u = null;
            if (obj != null)
            {
                u = new User()
                {
                    id = obj.Value<string>("id"),
                    text = obj.Value<string>("displayName"),
                    givenName = obj.Value<string>("givenName"),
                    surname = obj.Value<string>("surname"),
                    jobTitle = obj.Value<string>("jobTitle"),
                    mail = obj.Value<string>("mail"),
                    userPrincipalName = obj.Value<string>("userPrincipalName"),
                    mobilePhone = obj.Value<string>("mobilePhone"),
                    officeLocation = obj.Value<string>("officeLocation")
                };
            }

            return u;
        }


        /// <summary>
        /// Parses JArray to generic List of Group objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of Group objects</returns>
        public static List<Group> ToGroupList(this JArray array)
        {
            List<Group> groups = new List<Group>();
            foreach (var item in array)
                groups.Add(item.ToGroup());
            return groups;
        }

        /// <summary>
        /// Parses JToken to User object
        /// </summary>
        /// <param name="array">JToken</param>
        /// <returns>User object</returns>
        public static Group ToGroup(this JToken obj)
        {
            Group g = null;
            if (obj != null)
            {
                g = new Group()
                {
                    id = obj.Value<string>("id"),
                    text = obj.Value<string>("displayName"),
                    description = obj.Value<string>("description"),
                    mail = obj.Value<string>("mail"),
                    itemType = ItemType.Group,
                    visibility = obj.Value<string>("visibility")
                };
            }

            return g;
        }

        /// <summary>
        /// Parses JArray to generic List of Plan objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of Plan objects</returns>
        public static List<Plan> ToPlanList(this JArray array)
        {
            return array.Select(item => item.ToPlan()).ToList();
        }

        /// <summary>
        /// Parses JToken to Plan object
        /// </summary>
        /// <param name="token">JToken</param>
        /// <returns>Plan object</returns>
        public static Plan ToPlan(this JToken token)
        {
            return token.ToObject<Plan>();
        }

        /// <summary>
        /// Parses JArray to generic List of Bucket objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of Bucket objects</returns>
        public static List<Bucket> ToBucketList(this JArray array)
        {
            return array.Select(item => item.ToBucket()).ToList();
        }

        /// <summary>
        /// Parses JToken to Bucket object
        /// </summary>
        /// <param name="token">JToken</param>
        /// <returns>Bucket object</returns>
        public static Bucket ToBucket(this JToken token)
        {
            return token.ToObject<Bucket>();
        }

        /// <summary>
        /// Parses JArray to generic List of Task objects
        /// </summary>
        /// <param name="array">JArray</param>
        /// <returns>List of Task objects</returns>
        public static List<PlanTask> ToTasksList(this JArray array)
        {
            return array.Select(item => item.ToTask()).ToList();
        }

        /// <summary>
        /// Parses JToken to Task object
        /// </summary>
        /// <param name="token">JToken</param>
        /// <returns>Task object</returns>
        public static PlanTask ToTask(this JToken token)
        {
            return token.ToObject<PlanTask>();
        }
    }
}
