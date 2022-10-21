using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;

namespace SPS_Query_Tests.Models
{
    public class SPQueryLog
    {
        public int TotalObjectsReturned { get; set; }
        public bool HasErrors { get; set; }
        public List<Object> RequestData { get; set; }
    }

    public class SPRequestData
    {
        public IDictionary<string, object> QueryProperties { get; set; }
        public ClientResult<ResultTableCollection>? Query { get; set; }
    }
}
