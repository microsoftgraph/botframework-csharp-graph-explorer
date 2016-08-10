using System;
using Microsoft.Bot.Builder.Dialogs;
using MicrosoftGraphBot.Models;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class BucketLookupDialog : EntityLookupDialog<Bucket>
    {
        public BucketLookupDialog()
            : base("What bucket are you interested in (lookup by the bucket name)?",
                "No choices found... do you want to try again?",
                "Multiple choices found... which bucket would you like to pick?",
                GetRequestUri,
                Extensions.ToBucketList,
                (p, q) => p.Name.ToLower().Contains(q))
        {
        }

        private static string GetRequestUri(IDialogContext dialogContext)
        {
            var dialogEntity = dialogContext.ConversationData.GetDialogEntity();
            var id = dialogEntity.id;
            return $"https://graph.microsoft.com/beta/plans/{id}/buckets";
        }
    }
}