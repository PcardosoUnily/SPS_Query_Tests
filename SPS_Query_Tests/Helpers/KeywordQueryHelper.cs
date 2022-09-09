using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using SPS_Query_Tests.Constants;

namespace SPS_Query_Tests.Helpers
{
    public static class KeywordQueryHelper
    {
        public static KeywordQuery BuildQuery(ClientContext context, Dictionary<string, string> queryProperties)
        {
            var keywordQuery = new KeywordQuery(context)
            {
                SourceId = Guid.Parse(queryProperties[nameof(KeywordQuery.SourceId)]),
                QueryText = queryProperties[nameof(KeywordQuery.QueryText)],
                RowLimit = int.Parse(queryProperties[nameof(KeywordQuery.RowLimit)]),
                TrimDuplicates = bool.Parse(queryProperties[nameof(KeywordQuery.TrimDuplicates)]),
                TotalRowsExactMinimum = int.Parse(queryProperties[nameof(KeywordQuery.TotalRowsExactMinimum)]),
                EnableSorting = true
            };

            keywordQuery.SortList.Add(queryProperties[QueryConstants.SortListProperty], SortDirection.Ascending);
            keywordQuery.SelectProperties.Add(queryProperties[QueryConstants.PaginationProperty]);

            foreach (var prop in queryProperties[nameof(KeywordQuery.SelectProperties)].Split(','))
            {
                if (prop != QueryConstants.SortListProperty && prop != QueryConstants.PaginationProperty)
                {
                    keywordQuery.SelectProperties.Add(prop);
                }
            }

            return keywordQuery;
        }
    }
}
