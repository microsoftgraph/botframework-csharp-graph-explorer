using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using MicrosoftGraphBot.Models;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public abstract class EntityLookupDialog<T> : IDialog<T> 
    {
        private List<T> _entities;

        public abstract string LookupPrompt { get; }

        public abstract string NoLookupPrompt { get; }

        public abstract string NoChoicesPrompt { get; }

        public abstract string MultipleChoicesPrompt { get; }

        public abstract string GetRequestUri(IDialogContext dialogContext);

        public abstract List<T> DeserializeArray(JArray array);

        public abstract bool FilterEntity(T entity, string query);

#pragma warning disable 1998
        public async Task StartAsync(IDialogContext context)
#pragma warning restore 1998
        {
            context.Wait(ShowTextDialogAsync);
        }

#pragma warning disable 1998
        public async Task ShowTextDialogAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
#pragma warning restore 1998
        {
            // Get entities.
            var accessToken = await context.GetAccessToken();
            _entities = await GetEntitesAsync(context, accessToken);

            // No matches, retry.
            if (_entities.Count == 0)
            {
                Retry(context);
            }
            else
            {
                // If the entities are below five, let the user pick
                // one right away. If not, let the user do a lookup.
                if (_entities.Count <= 5)
                {
                    ShowChoices(context, _entities, true);
                }
                else
                {
                    PromptDialog.Text(context, OnTextDialogResumeAsync, LookupPrompt);
                }
            }
        }

        public async Task<List<T>> GetEntitesAsync(IDialogContext context, string accessToken)
        {
            // Create HTTP Client and get the response.
            var httpClient = new HttpClient();
            var requestUri = GetRequestUri(context);
            var json = await httpClient.MSGraphGET(accessToken, requestUri);

            // Deserialize the response.
            var response = DeserializeArray((JArray)json["value"]);
            return response;
        }

        private async Task OnTextDialogResumeAsync(IDialogContext context, IAwaitable<string> result)
        {
            // Filter the entities..
            var query = (await result).ToLower();
            var matches = _entities.Where(e => FilterEntity(e, query))
                .Take(5)
                .ToList();

            // Check the number of matches.
            switch (matches.Count)
            {
                case 0:
                    // No matches, retry.
                    Retry(context);
                    break;
                case 1:
                    // Resolve the exact match.
                    context.Done(matches[0]);
                    break;
                default:
                    ShowChoices(context, matches, false);
                    break;
            }
        }

        private void Retry(IDialogContext context)
        {
            PromptDialog.Confirm(context, async (retryContext, retryResult) =>
            {
                // Check retry result.
                var retry = await retryResult;
                if (retry)
                {
                    await ShowTextDialogAsync(retryContext, null);
                }
                else
                {
                    retryContext.Done<User>(null);
                }
            }, NoChoicesPrompt);
        }

        private void ShowChoices(IDialogContext context, IEnumerable<T> choices, bool firstPrompt)
        {
            PromptDialog.Choice(context, async (choiceContext, choiceResult) =>
            {
                var selection = await choiceResult;
                choiceContext.Done(selection);
            }, choices, firstPrompt ? NoLookupPrompt : MultipleChoicesPrompt);
        }
    }
}