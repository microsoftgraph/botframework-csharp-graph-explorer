using System;
using Newtonsoft.Json;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class PlanTask : ItemBase
    {
        private string _id;
        private string _title;

        [JsonProperty("@odata.etag")]
        public string ETag { get; set; }

        public string Id
        {
            get { return _id; }
            set
            {
                _id = value;
                navEndpoint = $"/tasks/{value}";
            }
        }

        public string PlanId { get; set; }

        public string BucketId { get; set; }

        public string Title
        {
            get { return _title; }
            set
            {
                _title = value;
                text = value;
            }
        }

        public string CreatedBy { get; set; }

        public string AssignedTo { get; set; }

        public string OrderHint { get; set; }

        public string AssigneePriority { get; set; }

        public int PercentComplete { get; set; }

        public string StartDateTime { get; set; }

        public string AssignedDateTime { get; set; }

        public string CreatedDateTime { get; set; }

        public string AssignedBy { get; set; }

        public string DueDateTime { get; set; }

        public string PreviewType { get; set; }

        public string CompletedDateTime { get; set; }

        public string ConversationThreadId { get; set; }

        public PlanTask()
            : base(null, null, ItemType.Task)
        {

        }
    }
}