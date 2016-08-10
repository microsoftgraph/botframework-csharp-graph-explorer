using System;
using Newtonsoft.Json;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class Bucket : ItemBase
    {
        [JsonProperty("@odata.etag")]
        public string ETag { get; set; }

        public string Id
        {
            get { return id; }
            set
            {
                id = value;
                navEndpoint = $"/tasks/{value}";
            }
        }

        public string Name
        {
            get { return text; }
            set
            {
                text = value;
            }
        }

        public string PlanId { get; set; }

        public string OrderHint { get; set; }

        public Bucket()
            : base(null, null, ItemType.Bucket)
        {

        }
    }
}