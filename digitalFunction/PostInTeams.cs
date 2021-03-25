using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Net.Http.Headers;
using RestSharp;

namespace digitalFunction
{
    public static class PostInTeams
    {
        private static GraphServiceClient _graphServiceClient;
        private static HttpClient httpClient = new HttpClient();
        [FunctionName("PostTeams")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            
            string requestBody = String.Empty;
            if (req.Body!=null)
            {
                string? message;
                using (StreamReader streamReader = new StreamReader(req.Body))
                {
                    requestBody = await streamReader.ReadToEndAsync();
                }
               
                dynamic data = JsonConvert.DeserializeObject(requestBody);
                message  =data?.message;
                message = message ?? "Hey this is hard coded message";
            }
            GraphServiceClient graphClient = GetAuthenticatedGraphClient();
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "100")
            };
            var graphResult = graphClient.Users.Request(options).GetAsync().Result;
            var chatMessage = new ChatMessage()
            {
                Body=new ItemBody { Content="Hi message"}
            };


            //await graphClient.Teams["efd7789e-4473-4473-967b-529ec6b2055e"].Channels["19:169c863c64894429802f36f7e0614273@thread.tacv2  "].Messages
              // .Request()
               //.AddAsync(chatMessage);
            var client = new RestClient("https://ctsmpn.webhook.office.com/webhookb2/efd7789e-4473-4473-967b-529ec6b2055e@525f3e23-ccc0-42c9-9fd7-ad446be35119/IncomingWebhook/375abf47d1ac44319d29a145364ef0e4/49ba6e73-6df7-441b-98be-8cd747f2c631");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("application/json", "{\r\n    \"@type\": \"MessageCard\",\r\n    \"@context\": \"http://schema.org/extensions\",\r\n    \"themeColor\": \"0076D7\",\r\n    \"summary\": \"Ramakrishnan created a new task\",\r\n    \"sections\": [{\r\n        \"activityTitle\": \"Hi Suresh! Ramakrishnan created a new task\",\r\n        \"activitySubtitle\": \"On Project Tango\",\r\n        \"activityImage\": \"https://teamsnodesample.azurewebsites.net/static/img/image5.png\",\r\n        \"facts\": [{\r\n            \"name\": \"Assigned to\",\r\n            \"value\": \"Unassigned\"\r\n        }, {\r\n            \"name\": \"Due date\",\r\n            \"value\": \"Mon May 01 2021 17:07:18 GMT+5:30 (Indian Standard Time)\"\r\n        }, {\r\n            \"name\": \"Status\",\r\n            \"value\": \"Not started\"\r\n        }],\r\n        \"markdown\": true\r\n    }],\r\n    \"potentialAction\": [{\r\n        \"@type\": \"ActionCard\",\r\n        \"name\": \"Add a comment\",\r\n        \"inputs\": [{\r\n            \"@type\": \"TextInput\",\r\n            \"id\": \"comment\",\r\n            \"isMultiline\": false,\r\n            \"title\": \"Add a comment here for this task\"\r\n        }],\r\n        \"actions\": [{\r\n            \"@type\": \"HttpPOST\",\r\n            \"name\": \"Add comment\",\r\n            \"target\": \"https://docs.microsoft.com/outlook/actionable-messages\"\r\n        }]\r\n    }, {\r\n        \"@type\": \"ActionCard\",\r\n        \"name\": \"Set due date\",\r\n        \"inputs\": [{\r\n            \"@type\": \"DateInput\",\r\n            \"id\": \"dueDate\",\r\n            \"title\": \"Enter a due date for this task\"\r\n        }],\r\n        \"actions\": [{\r\n            \"@type\": \"HttpPOST\",\r\n            \"name\": \"Save\",\r\n            \"target\": \"https://docs.microsoft.com/outlook/actionable-messages\"\r\n        }]\r\n    }, {\r\n        \"@type\": \"OpenUri\",\r\n        \"name\": \"Learn More\",\r\n        \"targets\": [{\r\n            \"os\": \"default\",\r\n            \"uri\": \"https://docs.microsoft.com/outlook/actionable-messages\"\r\n        }]\r\n    }, {\r\n        \"@type\": \"ActionCard\",\r\n        \"name\": \"Change status\",\r\n        \"inputs\": [{\r\n            \"@type\": \"MultichoiceInput\",\r\n            \"id\": \"list\",\r\n            \"title\": \"Select a status\",\r\n            \"isMultiSelect\": \"false\",\r\n            \"choices\": [{\r\n                \"display\": \"In Progress\",\r\n                \"value\": \"1\"\r\n            }, {\r\n                \"display\": \"Active\",\r\n                \"value\": \"2\"\r\n            }, {\r\n                \"display\": \"Closed\",\r\n                \"value\": \"3\"\r\n            }]\r\n        }],\r\n        \"actions\": [{\r\n            \"@type\": \"HttpPOST\",\r\n            \"name\": \"Save\",\r\n            \"target\": \"https://docs.microsoft.com/outlook/actionable-messages\"\r\n        }]\r\n    }]\r\n}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            // Read response content
            


            string responseMessage = graphResult != null ? graphResult[81].DisplayName : "null";


            return new OkObjectResult(responseMessage + "  " + response.StatusCode + " " + response.Content);
        }

        private static IAuthenticationProvider CreateAuthorizationProvider()
        {
            var clientId = System.Environment.GetEnvironmentVariable("AzureADAppClientId", EnvironmentVariableTarget.Process);
            var clientSecret = System.Environment.GetEnvironmentVariable("AzureADAppClientSecret", EnvironmentVariableTarget.Process);
            var redirectUri = System.Environment.GetEnvironmentVariable("AzureADAppRedirectUri", EnvironmentVariableTarget.Process);
            var tenantId = System.Environment.GetEnvironmentVariable("AzureADAppTenantId", EnvironmentVariableTarget.Process);
            var authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";

            //this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");


            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .WithRedirectUri(redirectUri)
                                              .WithClientSecret(clientSecret)
                                              .Build();

            return new MsalAuthenticationProvider(cca, scopes.ToArray()); ;
        }
        private static GraphServiceClient GetAuthenticatedGraphClient()
        {
            var authenticationProvider = CreateAuthorizationProvider();
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }
    }
    

     
}

