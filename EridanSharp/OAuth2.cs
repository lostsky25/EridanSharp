using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{
    public class OAuth2
    {
        private string redirectUri; //The address wich will open after the authentication.
        private string clientId;    //Client ID from console cloud google.
        private string clientSecret;//Client secret from console cloud google.
        public string accessToken; // ---.

        private string authorizationRequest; // The OAuth 2.0 authorization request.
        private const string requestAPI = "https://accounts.google.com/o/oauth2/v2/auth"; // The beginning of the authorizationRequest.
        private const string responseType = "code"; // The beginning of the authorizationRequest.
        private const string scope = "https://www.googleapis.com/auth/gmail.send"; // ---.
        private const string accessType = "offline"; // ---.
        private const string includeGrantedScopes = "true"; // ---.
        private const string state = "state_parameter_passthrough_value"; // ---.
        private const string grantType = "authorization_code"; // ---.
        HttpListenerContext context;

        public OAuth2(string clientId, string clientSecret)
        {
            this.clientId = clientId;
            this.clientSecret = clientSecret;
        }
        public async Task<bool> Authentication()
        {
            Debug.WriteLine("Start authentication.");
            // Creates a redirect URI using an available port on the loopback address.
            redirectUri = $"http://{IPAddress.Loopback}:{GetRandomUnusedPort()}/";
            Debug.WriteLine("Redirect URI: " + redirectUri);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectUri);
            Debug.WriteLine("Listening..");
            http.Start();
            // Creates the OAuth 2.0 authorization request.
            //client id: 271195115551-ceelr2teuhbeq3guse4edk13pcjapaea.apps.googleusercontent.com
            string authorizationRequest =
                requestAPI +
                "?response_type=" + responseType +
                "&scope=" + scope +
                "&access_type=" + accessType +
                "&include_granted_scopes=" + includeGrantedScopes +
                "&state=" + state +
                "&client_id=" + clientId +
                "&redirect_uri=" + redirectUri;

            // Opens request in the browser.
            Process.Start(authorizationRequest);

            // Waits for the OAuth authorization response.

            context = await http.GetContextAsync();

            // Makes token request.
            string tokenBody =
                "code=" + GetContextValue("code") +
                "&client_id=" + clientId +
                "&client_secret=" + clientSecret +
                "&redirect_uri=" + redirectUri +
                "&grant_type=" + grantType;

            // Sends the request
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create("https://accounts.google.com/o/oauth2/token");
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            byte[] tokenRequestBodyBytes = Encoding.ASCII.GetBytes(tokenBody);
            tokenRequest.ContentLength = tokenRequestBodyBytes.Length;
            using (Stream requestStream = tokenRequest.GetRequestStream())
            {
                await requestStream.WriteAsync(tokenRequestBodyBytes, 0, tokenRequestBodyBytes.Length);
            }

            try
            {
                // gets the response
                WebResponse tokenResponse = await tokenRequest.GetResponseAsync();
                using (StreamReader reader = new StreamReader(tokenResponse.GetResponseStream()))
                {
                    // reads response body
                    string responseText = await reader.ReadToEndAsync();
                    Debug.WriteLine(responseText);

                    // converts to dictionary
                    Dictionary<string, string> tokenEndpointDecoded = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseText);

                    accessToken = tokenEndpointDecoded["access_token"];
                    //await RequestUserInfoAsync(accessToken);
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        Debug.WriteLine("HTTP: " + response.StatusCode);
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                        {
                            // reads response body
                            string responseText = await reader.ReadToEndAsync();
                            Debug.WriteLine(responseText);
                        }
                    }
                }
            }

            return true;
        }

        public async Task<int> send(MimeMessage message)
        {
            var url = "https://www.googleapis.com/gmail/v1/users/drsail043@gmail.com/messages/send";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";

            httpRequest.Accept = "application/json";
            httpRequest.ContentType = "application/json";
            httpRequest.Headers["Authorization"] = "Bearer " + accessToken;

            string data = "{\"raw\":" + $" \"{message.GetMessageBase64()}" + @"""}";

            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(data);
            }

            var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
            }

            Debug.WriteLine(httpResponse.StatusCode);

            return 0;
        }

        public string GetContextValue(string name)
        {
            return context.Request.QueryString.Get(name);
        }

        public static int GetRandomUnusedPort()
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }
    }
}
