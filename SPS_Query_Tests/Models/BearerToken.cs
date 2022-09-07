using Newtonsoft.Json;

namespace SPS_Query_Tests.Models
{
    public class BearerToken
    {
        [JsonProperty("resource")]
        public string Resource { get; set; }
        [JsonProperty("expires_in")]
        public string Expires_In { get; set; }
        [JsonProperty("access_token")]
        public string Access_Token { get; set; }
        public DateTime Expiration_On { get; set; }
    }
}
