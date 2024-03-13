using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    public class CreateFromMSTeamsResult
    {
        [JsonPropertyName("url")]
        public string Url { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("number")]
        public string Number { get; set; }
    }
}
