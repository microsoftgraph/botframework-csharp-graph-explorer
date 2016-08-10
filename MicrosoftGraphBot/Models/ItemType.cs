using System;

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
        Bucket,
        Task
    }
}
