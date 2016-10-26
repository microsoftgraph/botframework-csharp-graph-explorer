using System;
using System.Collections.Generic;
using Microsoft.Bot.Builder.Dialogs;
using MicrosoftGraphBot.Models;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class BucketLookupDialog : EntityLookupDialog<Bucket>
    {
        public override string LookupPrompt => "Which bucket are you interested in (lookup by full/partial name)?";

        public override string NoLookupPrompt => "Which bucket are you interested in?";

        public override string NoChoicesPrompt => "No choices found... do you want to try again?";

        public override string MultipleChoicesPrompt => "Multiple choices found... which bucket would you like to choose?";

        public override string GetRequestUri(IDialogContext dialogContext)
        {
            var plan = dialogContext.ConversationData.Get<Plan>("Plan");
            return $"https://graph.microsoft.com/beta/plans/{plan.id}/buckets";
        }

        public override List<Bucket> DeserializeArray(JArray array)
        {
            return array.ToBucketList();
        }

        public override bool FilterEntity(Bucket entity, string query)
        {
            return entity.Name.ToLower().Contains(query);
        }
    }
}