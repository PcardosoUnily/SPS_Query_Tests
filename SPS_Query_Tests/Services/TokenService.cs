using RestSharp;
using SPS_Query_Tests.Models;
using System.Web;

namespace SPS_Query_Tests.Services
{
    public static class TokenService
    {
        public const string Authority = "https://accounts.accesscontrol.windows.net";
        public const string ResourceGuid = "00000003-0000-0ff1-ce00-000000000000";

        /// <remarks>See https://anexinet.com/blog/getting-an-access-token-for-sharepoint-online/ for details</remarks>
        public static BearerToken GetNewAppContextToken(string tenantId, string clientId, string clientSecret, Uri uri)
        {
            var client = new RestClient($"{Authority}/{HttpUtility.UrlEncode(tenantId)}/tokens/OAuth/2");

            var request = new RestRequest();
            request.AddParameter("grant_type", "client_credentials");
            request.AddParameter("client_id", $"{clientId}@{tenantId}");
            request.AddParameter("client_secret", clientSecret);
            request.AddParameter("resource", $"{ResourceGuid}/{uri.Host}@{tenantId}");

            var response = client.Post<BearerToken>(request);
            return response;
        }
    }
}
