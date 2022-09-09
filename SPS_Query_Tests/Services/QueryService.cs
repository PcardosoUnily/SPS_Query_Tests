using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using SPS_Query_Tests.Constants;
using SPS_Query_Tests.Helpers;

namespace SPS_Query_Tests.Services
{
    public static class QueryService
    {
        public static void RunSearchQuery(ClientContext context, Dictionary<string, string> queryParameters)
        {
            var correlationId = string.Empty;
            var initialQueryText = queryParameters[QueryConstants.QueryTextProperty];
            var paginationProperty = queryParameters[QueryConstants.PaginationProperty];
            var lastDocId = string.Empty;
            int rowsRetrieved = 0;
            int totalRows = 1;
            ResultTable results;

            // For some reason, the MS Docs have us set the SortList property value as [DocId] but that's not the actual key value of the property returned
            // We must thus sanitize this value to make sure we can retrieve the property from the results
            // This is quite janky, once again refer to: https://docs.microsoft.com/en-us/sharepoint/dev/general-development/pagination-for-large-result-sets
            // See confirmation: https://github.com/SharePoint/sp-dev-docs/issues/8426
            var sanitizedPaginationProp = queryParameters[QueryConstants.SortListProperty].Trim('[', ']').Trim();

            do
            {
                try
                {
                    var executor = new SearchExecutor(context);
                    var query = KeywordQueryHelper.BuildQuery(context, queryParameters);
                    if (!string.IsNullOrEmpty(lastDocId))
                    {
                        query.QueryText = $"{paginationProperty}>{lastDocId} AND {initialQueryText}";
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
                        totalRows = results.TotalRows;
                        rowsRetrieved += results.RowCount;

                        if (totalRows != 0)
                        {
                            lastDocId = results.ResultRows.Last()[sanitizedPaginationProp].ToString();
                        }

                        Console.WriteLine($"{rowsRetrieved} Total Rows retrieved. {totalRows} Rows left to retrieve. CorrelationId: {correlationId}");
                    }
                    else
                    {
                        Console.WriteLine($"Sharepoint Search returned a Null response with no data! Current Context Correlation Id: {context.TraceCorrelationId}");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception thrown while executing SPS Query Request!");
                    Console.WriteLine(ex.ToString());

                    break;
                }

            } while (totalRows != 0);
        }
    }
}
