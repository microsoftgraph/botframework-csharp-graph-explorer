using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public enum OperationType
    {
        Manager,
        DirectReports,
        Photo,
        Files,
        Mail,
        Events,
        Contacts,
        Groups,
        WorkingWith,
        TrendingAround,
        People,
        Notebooks,
        Tasks,
        Plans,
        Members,
        Conversations,

        //Navigation choices
        Next,
        Previous,
        Up,
        StartOver,
        ChangeDialogEntity,
        ShowOperations,
        Create,
        Delete,
        Download,
        Upload,
        Folder,
        InProgress,
        Complete
    }
}
