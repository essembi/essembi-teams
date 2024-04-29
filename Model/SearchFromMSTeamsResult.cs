using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    public class SearchFromMSTeamsResult
    {
        [JsonPropertyName("url")]
        public string Url { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("table")]
        public string Table { get; set; }
    }

    public class SearchFromMSTeamsResults
    {
        [JsonPropertyName("results")]
        public SearchFromMSTeamsResult[] Results { get; set; }
    }
}
