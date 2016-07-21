using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class Mail : ItemBase
    {
        public Mail() : base() { }
        public Mail(string text, string endpoint, ItemType type) : base(text, endpoint, type) { }

        public string importance { get; set; }
        public string senderName { get; set; }
        public string senderEmail { get; set; }
        public DateTime sentDate { get; set; }
        public bool isRead { get; set; }
    }
}
