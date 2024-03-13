using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    public class MessageResponse
    {
        [JsonPropertyName("message")]
        public string Message { get; set; }
    }
}
