using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public enum ItemType
    {
        File,
        Folder,
        Mail,
        Event,
        Contact,
        Person,
        Group,
        NavNext,
        NavPrevious,
        NavUp,
        Cancel,
        Plan,
        Task,
    }
}
