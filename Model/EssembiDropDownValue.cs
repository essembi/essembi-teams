using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    internal class EssembiDropDownValue
    {
        [JsonPropertyName("id")]
        public long Id { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }
    }
}
