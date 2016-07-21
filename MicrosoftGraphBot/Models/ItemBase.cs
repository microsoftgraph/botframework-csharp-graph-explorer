using MicrosoftGraphBot.Dialog;
using Microsoft.Bot.Builder.Dialogs;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class ItemBase
    {
        public ItemBase() { }
        public ItemBase(string text, string endpoint, ItemType type)
        {
            this.text = text;
            this.navEndpoint = endpoint;
            this.itemType = type;
        }

        public string id { get; set; }
        public string text { get; set; }
        public string navEndpoint { get; set; }
        public ItemType itemType { get; set; }

        public override string ToString()
        {
            return this.text;
        }
    }
}
