using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Newtonsoft.Json;
using SPS_Query_Tests.Constants;
using SPS_Query_Tests.Helpers;
using SPS_Query_Tests.Models;
using File = System.IO.File;

namespace SPS_Query_Tests.Services
{
    public static class QueryService
    {
        public static void RunSearchQuery(ClientContext context, Dictionary<string, string> queryParameters, string exportPath)
        {
            var correlationId = string.Empty;
            var initialQueryText = queryParameters[QueryConstants.QueryTextProperty];
            var paginationProperty = queryParameters[QueryConstants.PaginationProperty];
            var lastDocId = string.Empty;
            int totalRows = 1;
            ResultTable results;
            Guid queryId = Guid.NewGuid();
            SPQueryLog spRequest = new()
            {
                TotalObjectsReturned = 0,
                RequestData = new()
            };

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
                    ClientResult<ResultTableCollection>? response = executor.ExecuteQuery(query);
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

                        // Build SPResponse object and add to List
                        // This serves as a way to store the raw JSON data from the response itself
                        // We can capture any relevant response data this way, instead of relying on more manual methods like Fiddler
                        // ClientResult<ResultTableCollection> contains duplicate Key "Properties" which is not serialized by JsonConvert as it's technically not valid json
                        // We thus create a new Class "SPResponseData" and assign the 1st Properties data to a "QueryProperties" key instead
                        SPRequestData responseData = new()
                        {
                            QueryProperties = response?.Value?.Properties ?? null,
                            Query = response
                        };

                        spRequest.RequestData.Add(responseData);

                        // Get RelevantResults table data to output how many results we got
                        results = response.Value.Single(table => table.TableType == "RelevantResults");
                        totalRows = results.TotalRows;
                        spRequest.TotalObjectsReturned += results.RowCount;

                        if (totalRows != 0)
                        {
                            lastDocId = results.ResultRows.Last()[sanitizedPaginationProp].ToString();
                        }

                        Console.WriteLine($"{spRequest.TotalObjectsReturned} Total Rows retrieved. {totalRows} Rows left to retrieve. CorrelationId: {correlationId}");
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

                    spRequest.HasErrors = true;
                    spRequest.RequestData.Add(ex);

                    break;
                }
            } 
            while (totalRows != 0);

            string path = $"{exportPath.TrimEnd('/').Trim()}\\{queryId}.json";
            Console.WriteLine($"Exporting Query data to {path}");
            var jsonData = JsonConvert.SerializeObject(spRequest, Formatting.Indented);
            File.WriteAllText(path, jsonData);
        }
    }
}
