using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using SPS_Query_Tests.Helpers;

namespace SPS_Query_Tests.Services
{
    public static class QueryService
    {
        private const string QueryTextProperty = "QueryText";
        private const string PaginationProperty = "PaginationProperty";

        public static void RunSearchQuery(ClientContext context, Dictionary<string, string> queryParameters)
        {
            var initialQueryText = queryParameters[QueryTextProperty];
            var paginationProperty = queryParameters[PaginationProperty];
            var lastIndexDocId = string.Empty;
            int rowsRetrieved = 0;
            int rowsToRetrieve = 1;
            int totalRows = 0;
            ResultTable results;

            var correlationId = string.Empty;
            while (rowsToRetrieve != 0)
            {
                try
                {
                    var executor = new SearchExecutor(context);
                    var query = KeywordQueryHelper.BuildQuery(context, queryParameters);
                    if (!string.IsNullOrEmpty(lastIndexDocId))
                    {
                        query.QueryText = $"{initialQueryText} AND {paginationProperty}>{lastIndexDocId}";
                    }

                    Console.WriteLine($"Executing SP query. Query text: {query.QueryText}");
                    var response = executor.ExecuteQuery(query);
                    context.ExecuteQuery();

                    if (response != null)
                    {
                        if (response.Value.Properties.TryGetValue("CorrelationId", out var correlationIdLatest))
                        {
                            correlationId = correlationIdLatest.ToString();
                        }
                        else
                        {
                            correlationId = context.TraceCorrelationId;
                        }

                        results = response.Value.Single(table => table.TableType == "RelevantResults");
                        lastIndexDocId = results.ResultRows.Last()[paginationProperty].ToString();

                        rowsRetrieved += results.RowCount;
                        totalRows = results.TotalRows;
                        rowsToRetrieve = results.TotalRows;

                        Console.WriteLine($"{rowsRetrieved} Total Rows retrieved. {totalRows} Rows left to retrieve. CorrelationId: {correlationId}");
                    }
                    else
                    {
                        Console.WriteLine($"Sharepoint Search returned a Null response with no data! Current Context Correlation Id: {context.TraceCorrelationId}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception thrown while executing SPS Query Request!");
                    Console.WriteLine(ex.ToString());

                    break;
                }

            }
        }
        private static string ToStringNullSafe(this object value)
        {
            return (value ?? string.Empty).ToString();
        }
    }
}
