using AuthBot;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using MicrosoftGraphBot.Dialog;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis.Models;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class Group : ItemBase
    {
        public Group() : base() { }
        public Group(string text, string endpoint, ItemType type) : base(text, endpoint, type) { }
        public string description { get; set; }
        public string mail { get; set; }
        public string[] groupTypes { get; set; }
        public string visibility { get; set; }

        public override string ToString()
        {
            return this.text;
        }

        public static async Task<List<Group>> Lookup(string token, string searchPhrase)
        {
            List<Group> groups = new List<Group>();
            HttpClient client = new HttpClient();

            //search for the user
            var endpoint = String.Format("https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,'{0}')%20or%20startswith(mail,'{0}')", searchPhrase);
            var json = await client.MSGraphGET(token, endpoint);
            groups = ((JArray)json["value"]).ToGroupList();

            return groups;
        }
    }
}
