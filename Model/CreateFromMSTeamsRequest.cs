using System.Collections.Generic;

namespace Essembi.Integrations.Teams.Model
{
    public class CreateFromMSTeamsRequest
    {
        public string Email { get; set; }

        public long AppId { get; set; }

        public long TableId { get; set; }

        public Dictionary<string, object> Values { get; set; }
    }
}
