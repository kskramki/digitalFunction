using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Net.Http;
using System.Net.Http.Headers;

namespace digitalFunction
{

    public class MsalAuthenticationProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication _clientApplication;
    private string[] _scopes;

    public MsalAuthenticationProvider(IConfidentialClientApplication clientApplication, string[] scopes)
    {
        _clientApplication = clientApplication;
        _scopes = scopes;
    }

    /// <summary>
    /// Update HttpRequestMessage with credentials
    /// </summary>
    public async Task AuthenticateRequestAsync(HttpRequestMessage request)
    {
        var token = await GetTokenAsync();
        request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
    }

    /// <summary>
    /// Acquire Token 
    /// </summary>
    public async Task<string> GetTokenAsync()
    {
        AuthenticationResult authResult = null;
        authResult = await _clientApplication.AcquireTokenForClient(_scopes)
                            .ExecuteAsync();
        return authResult.AccessToken;
    }

     
    }
}
