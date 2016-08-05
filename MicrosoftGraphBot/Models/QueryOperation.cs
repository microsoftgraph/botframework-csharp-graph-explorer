using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class QueryOperation
    {
        public string Text { get; set; }
        public OperationType Type { get; set; }
        public string Endpoint { get; set; }
        public object ContextObject { get; set; }

        public override string ToString()
        {
            return Text;
        }

        public T GetContextObjectAs<T>() where T : class
        {
            if (ContextObject is T)
            {
                return (T)ContextObject;
            }
            if (ContextObject is JObject)
            {
                return ((JObject)ContextObject).ToObject<T>();
            }
            return null;
        }

        public static List<QueryOperation> GetEntityResourceTypes(EntityType entityType)
        {
            if (entityType == EntityType.Me)
                return meResourceTypes.ToQueryOperations();
            else if (entityType == EntityType.User)
                return userResourceTypes.ToQueryOperations();
            else if (entityType == EntityType.Group)
                return groupResourceTypes.ToQueryOperations();
            else
                return null;
        }

        private static List<OperationType> meResourceTypes = new List<OperationType>() {
            OperationType.Manager,
            OperationType.DirectReports,
            OperationType.Photo,
            OperationType.Files,
            OperationType.Mail,
            OperationType.Events,
            OperationType.Contacts,
            OperationType.Groups,
            OperationType.WorkingWith,
            OperationType.TrendingAround,
            OperationType.People,
            OperationType.Notebooks,
            OperationType.Tasks,
            OperationType.Plans
        };
        private static List<OperationType> userResourceTypes = new List<OperationType>() {
            OperationType.Manager,
            OperationType.DirectReports,
            OperationType.Photo,
            OperationType.Files,
            OperationType.Groups,
            OperationType.WorkingWith,
            OperationType.TrendingAround,
            OperationType.People,
            OperationType.Notebooks,
            OperationType.Tasks,
            OperationType.Plans
        };
        private static List<OperationType> groupResourceTypes = new List<OperationType>() {
            OperationType.Members,
            OperationType.Files,
            OperationType.Conversations,
            OperationType.Events,
            OperationType.Photo,
            OperationType.Notebooks
        };
    }
}
