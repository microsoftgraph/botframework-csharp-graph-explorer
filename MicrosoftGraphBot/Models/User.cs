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
    public class User : ItemBase
    {
        public User() : base() { }
        public User(string text, string endpoint, ItemType type) : base(text, endpoint, type) { }
        public bool isMe { get; set; }
        public string givenName { get; set; }
        public string surname { get; set; }
        public string jobTitle { get; set; }
        public string mail { get; set; }
        public string userPrincipalName { get; set; }
        public string mobilePhone { get; set; }
        public string officeLocation { get; set; }

        public override string ToString()
        {
            return this.text;
        }


        public static async Task<User> Me(string token)
        {
            HttpClient client = new HttpClient();

            //return the current user
            var json = await client.MSGraphGET(token, "https://graph.microsoft.com/v1.0/me");
            return json.ToUser();
        }

        public static async Task<List<User>> Lookup(string token, string searchPhrase)
        {
            List<User> users = new List<User>();
            HttpClient client = new HttpClient();

            //search for the user
            var endpoint = String.Format("https://graph.microsoft.com/v1.0/users?$filter=startswith(givenName,'{0}')%20or%20startswith(surname,'{0}')%20or%20startswith(displayName,'{0}')%20or%20startswith(userPrincipalName,'{0}')", searchPhrase);
            var json = await client.MSGraphGET(token, endpoint);
            users = ((JArray)json["value"]).ToUserList();

            return users;
        }
    }
}
