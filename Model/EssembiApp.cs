using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    internal class EssembiApp
    {
        [JsonPropertyName("id")]
        public long Id { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("tableId")]
        public long TableId { get; set; }

        [JsonPropertyName("fields")]
        public EssembiField[] Fields { get; set; }
    }
}
