using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using AuditLogAPICalls;
using cfg;
using System.Reflection;
using System.IO;
using System.Text.Json;

namespace AuditLogAPICalls
{
    class ApiCalls
    {
        // authentication credentials through the AAD app
        readonly string grantType = "client_credentials";
        string? tenantId;
        string? clientId;
        string? clientSecret;
        readonly string scope = "https://manage.office.com/.default";
        string startTime = DateTime.Now.AddDays(-1).ToString($"yyyy-MM-ddT00\\%3A00"); // yesterday 00:00 - can use a specific HH:mm instead of 00:00
        string endTime = DateTime.Now.AddDays(-1).ToString($"yyyy-MM-ddT23\\%3A59"); // yesterday 23:59 - same

        public void setClientId(string clientID)
        {
            this.clientId = clientID;
        }

        public void setClientSecret(string clientSecret)
        {
            this.clientSecret = clientSecret;
        }

        public void setTenantId(string tenantId)
        {
            this.tenantId = tenantId;
        }

        public async Task<string> APICallHelper(string accessToken, HttpMethod method, string requestUri)
        {
            var client = new HttpClient();
            var request = new HttpRequestMessage
            {
                Method = method, // HttpsMethod.Post or HttpMethod.Get
                RequestUri = new Uri(requestUri),
                Headers =
    {
        { "Authorization", $"Bearer {accessToken}" },
    },
            };
            using (var response = await client.SendAsync(request))
            {
                response.EnsureSuccessStatusCode();
                var body = await response.Content.ReadAsStringAsync();
                return body;
            }
        }

        public async Task<string> authTokenAPI()
        {
            // this call is a bit different, it cannot be done with the APICallHelper method
            var clientHandler = new HttpClientHandler();
            var client = new HttpClient(clientHandler);
            var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token"), // the new Uri() instance is not necessery, can be only string 
                Content = new FormUrlEncodedContent(new Dictionary<string, string>
    {
        { "grant_type", grantType },
        { "client_id", clientId },
        { "client_secret", clientSecret },
        { "scope", scope },
    }),
            };
            using (var response = await client.SendAsync(request))
            {
                response.EnsureSuccessStatusCode();
                var body = await response.Content.ReadAsStringAsync();
                //Console.WriteLine(body);
                var token = JObject.Parse(body)["access_token"].Value<string>();
                return token;
            }
        }

        public async Task<string> subscriptionStartAPI(string accessToken)
        {
            var body = await APICallHelper(accessToken, HttpMethod.Post, $"https://manage.office.com/api/v1.0/{tenantId}/activity/feed/subscriptions/start?contentType=Audit.SharePoint");
            //Console.WriteLine(body);
            return "Subscription started";
        }

        public async Task<List<object>> chunksGetterAPI(string accessToken)
        {
            var body = await APICallHelper(accessToken, HttpMethod.Get, $"https://manage.office.com/api/v1.0/{tenantId}/activity/feed/subscriptions/content?contentType=Audit.SharePoint&startTime={startTime}&endTime={endTime}");
            List<object> chunks = JsonConvert.DeserializeObject<List<object>>(body);
            return chunks; // the api call returns the logs bundled into chunks
        }
        public async Task<List<object>> contentGetterAPI(string accessToken, List<object> chunks)
        {
            var logs = new List<object>();

            // getting the contents from each chunk
            foreach (var chunk in chunks)
            {
                var contentUri = ((JObject)chunk)["contentUri"].Value<string>();
                var body = await APICallHelper(accessToken, HttpMethod.Get, contentUri);
                var contents = JsonConvert.DeserializeObject<List<object>>(body);

                foreach (var content in contents)
                {
                    //Console.WriteLine(content);
                    var log = new
                    {
                        CreationTime = ((JObject)content)["CreationTime"].Value<string>(),
                        Operation = ((JObject)content)["Operation"].Value<string>(),
                        UserId = ((JObject)content)["UserId"].Value<string>(),
                        ObjectId = ((JObject)content)["ObjectId"].Value<string>()
                    };

                    logs.Add(log);
                }
            }

            //foreach (var log in logs)
            //{
            //    Console.WriteLine(log);
            //    Console.WriteLine();
            //}

            return logs;
        }

        public async Task<string> subscriptionStopAPI(string accessToken)
        {
            var body = await APICallHelper(accessToken, HttpMethod.Post, $"https://manage.office.com/api/v1.0/{tenantId}/activity/feed/subscriptions/stop?contentType=Audit.SharePoint");
            //Console.WriteLine(body);
            return "Subscription ended";
        }
    }
}

class Program
{
    static async Task Main(string[] args)
    {
        // creating a new class instance
        var AuditlogRetrievingCalls = new ApiCalls();

        // setting the credentials from the config.cs
        var MY_CONFIG = new Config();
        AuditlogRetrievingCalls.setClientId(MY_CONFIG.CLIENT_ID);
        AuditlogRetrievingCalls.setClientSecret(MY_CONFIG.CLIENT_SECRET);
        AuditlogRetrievingCalls.setTenantId(MY_CONFIG.TENANT_ID);

        // getting the access token
        var accessToken = await AuditlogRetrievingCalls.authTokenAPI(); // always gets invoked because of the "await"
        //Console.WriteLine(accessToken);

        // starting subscription
        var subsStartResp = await AuditlogRetrievingCalls.subscriptionStartAPI(accessToken);
        Console.WriteLine(subsStartResp);

        // retrieving audit log chunks from sharepoint
        var logGetter = await AuditlogRetrievingCalls.chunksGetterAPI(accessToken);

        // getting content out of the logs
        var contentGetter = await AuditlogRetrievingCalls.contentGetterAPI(accessToken, logGetter);

        // serialize the result object to a JSON object, and save it locally
        var json = System.Text.Json.JsonSerializer.Serialize(contentGetter);
        string date = DateTime.Now.AddDays(-1).ToString($"yyyy-MM-dd");
        var path = @$"log_{date}.json";
        File.WriteAllText(path, json);

        // stopping subscription
        var subsStopResp = await AuditlogRetrievingCalls.subscriptionStopAPI(accessToken);
        Console.WriteLine(subsStopResp);
    }
}