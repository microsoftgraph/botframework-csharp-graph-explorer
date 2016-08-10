using System;
using Newtonsoft.Json;

namespace MicrosoftGraphBot.Models
{
    [Serializable]
    public class Plan : ItemBase
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

        public string Title
        {
            get { return text; }
            set
            {
                text = value;
            }
        }

        public string Owner { get; set; }

        public string CreatedBy { get; set; }

        public Plan()
            : base(null, null, ItemType.Plan)
        {

        }
    }
}