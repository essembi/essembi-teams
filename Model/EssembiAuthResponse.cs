using System.Text.Json.Serialization;

namespace Essembi.Integrations.Teams.Model
{
    internal class EssembiAuthResponse
    {
        [JsonPropertyName("apps")]
        public EssembiApp[] Apps { get; set; }
    }
}
