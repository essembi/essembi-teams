using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    internal class EssembiField
    {
        [JsonPropertyName("id")]
        public long Id { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("required")]
        public bool Required { get; set; }

        [JsonPropertyName("values")]
        public EssembiDropDownValue[] Values { get; set; }
    }
}
