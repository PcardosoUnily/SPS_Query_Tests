using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;

namespace SPS_Query_Tests.Helpers
{
    public static class KeywordQueryHelper
    {
        private const string SortListProperty = "SortListProperty";
        private const string PaginationProperty = "PaginationProperty";

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

            keywordQuery.SortList.Add(queryProperties[SortListProperty], SortDirection.Ascending);
            keywordQuery.SelectProperties.Add(queryProperties[PaginationProperty]);

            foreach (var prop in queryProperties[nameof(KeywordQuery.SelectProperties)].Split(','))
            {
                if (prop != SortListProperty && prop != PaginationProperty)
                {
                    keywordQuery.SelectProperties.Add(prop);
                }
            }

            return keywordQuery;
        }
    }
}
