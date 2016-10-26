using System;
using System.Collections.Generic;
using Microsoft.Bot.Builder.Dialogs;
using MicrosoftGraphBot.Models;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphBot.Dialog
{
    [Serializable]
    public class PlanLookupDialog : EntityLookupDialog<Plan>
    {
        public override string LookupPrompt => "Which plan are you interested in (lookup by full/partial name)?";

        public override string NoLookupPrompt => "Which plan are you interested in?";

        public override string NoChoicesPrompt => "No choices found... do you want to try again?";

        public override string MultipleChoicesPrompt => "Multiple choices found... which plan would you like to choose?";

        public override string GetRequestUri(IDialogContext dialogContext)
        {
            var user = dialogContext.ConversationData.Get<User>("Me");
            return $"https://graph.microsoft.com/beta/users/{user.id}/plans";
        }

        public override List<Plan> DeserializeArray(JArray array)
        {
            return array.ToPlanList();
        }

        public override bool FilterEntity(Plan entity, string query)
        {
            return entity.Title.ToLower().Contains(query);
        }
    }
}