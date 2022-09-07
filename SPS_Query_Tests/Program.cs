using Microsoft.SharePoint.Client;
using SPS_Query_Tests.Services;

// Define Authentication Variable Values
const string tenantId = ""; // Azure Tenant Id of SPO Environment (GUID)
const string clientId = ""; // SPO Service Principal Id (GUID)
const string clientSecret = ""; // SPO Service Principal Secret
const string siteUri = ""; // Site Uri
const string resultSourceId = ""; // Sharepoint Result Source GUID used for the Query

// Define Query Parameters
// If using docId for pagination, make sure to include IndexDocId in selected properties
// see: https://docs.microsoft.com/en-us/sharepoint/dev/general-development/pagination-for-large-result-sets
Dictionary<string, string> queryParameters = new()
{
    { "QueryText", "*" },
    { "SourceId", resultSourceId },
    { "SortListProperty", "[DocId]" },
    { "RowLimit", "500" },
    { "TotalRowsExactMinimum", "100" },
    { "TrimDuplicates", "true" },
    { "PaginationProperty", "IndexDocId" },
    { "SelectProperties", "UserProfile_GUID, AccountName" }
};

// Initialize ClientContext and retrieve Bearer Token
var context = new ClientContext(siteUri);
var token = TokenService.GetNewAppContextToken(tenantId, clientId, clientSecret, new Uri(siteUri));

// Add token to request event
// Feel free to change UserAgent value to something more identifying to yourself
context.ExecutingWebRequest += (sender, e) =>
{
    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token.Access_Token;
    e.WebRequestExecutor.WebRequest.UserAgent = "ISV|Unily|SPSQueryTests/1.0";
};

// Test context connection
context.Load(context.Web, x => x.Title);
context.ExecuteQuery();
Console.WriteLine($"Connected to {context.Web.Title}");
Console.WriteLine();

// Set how many time you want the configured query to run for
int loops = 1;
for (int i = 1; i <= loops; i++)
{
    Console.WriteLine("==================================");
    Console.WriteLine($"Running Query Loop: {i}");
    Console.WriteLine();

    // Execute Query
    QueryService.RunSearchQuery(context, queryParameters);

    Console.WriteLine();
    Console.WriteLine($"Finished Running Query Loop: {i}");
    Console.WriteLine("==================================");
    Console.WriteLine();
}

Console.WriteLine("Application Finished");
Console.WriteLine("Press any key to exit");
Console.ReadLine();
