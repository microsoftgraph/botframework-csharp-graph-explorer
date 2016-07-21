using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class File : ItemBase
    {
        public File() : base() { }
        public File(string text, string endpoint, ItemType type) : base(text, endpoint, type) { }

        public int size { get; set; }
        public string webUrl { get; set; }
    }
}
