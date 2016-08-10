using System;
using Microsoft.Bot.Builder.Dialogs;
using MicrosoftGraphBot.Models;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class PlanLookupDialog : EntityLookupDialog<Plan>
    {
        public PlanLookupDialog()
            : base("What plan are you interested in (lookup by the plan name)?",
                "No choices found... do you want to try again?",
                "Multiple choices found... which plan would you like to pick?",
                GetRequestUri,
                Extensions.ToPlanList,
                (p, q) => p.Title.ToLower().Contains(q))
        {
        }

        private static string GetRequestUri(IDialogContext dialogContext)
        {
            var dialogEntity = dialogContext.ConversationData.GetDialogEntity();
            var id = dialogEntity.id;
            return $"https://graph.microsoft.com/beta/users/{id}/plans";
        }
    }
}