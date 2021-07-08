using Newtonsoft.Json;
using SsPvo.Client.Messages.Base;

namespace SsPvo.Client.Models
{
    public class QueueMessagesResponse : IJsonResponse
    {
        [JsonProperty("messages")]
        public int Messages { get; set; }

        [JsonProperty("idJwts", Required = Required.Default)]
        public uint[] IdJwts { get; set; }
    }
}
