using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace EridanSharp
{
    public class OAuth2
    {
        private HttpListenerContext context;
        private OAuth2Info oAuth2Info;
        private OAuth2Request oAuth2Request;
        private OAuth2Profile oAuth2Profile;
        private HttpReq httpReq;
        private string pathToken;

        private class OAuth2Info
        {
            public string access_token;          // Access token is the thing that applications use to make API requests on behalf of a user.
            //public string authorizationCode;    // The authorization code grant is used when an application exchanges an authorization code for an access token.
            public string redirectUri;          // The address wich will open after the authentication.
            public string clientId;             // Client ID from console cloud google.
            public string clientSecret;         // Client secret from console cloud google.
            public string sucessPage;           // The offline page will open after successful authorization.
            public string unsucessPage;         // The offline page will open after unsuccessful authorization.
            public string expires_in;            // How long to live\available current token.
            public string refresh_token;         // Needs for updating current token.
        }

        private class OAuth2Request
        {
            public readonly string requestAPI = "https://accounts.google.com/o/oauth2/v2/auth";   // The beginning part of the authorization request.
            public readonly string scope = "https://www.googleapis.com/auth/gmail.modify";          // Scope is a mechanism in OAuth 2.0 to limit an application's access to a user's account.
            public readonly string responseType = "code";                                         // The Response Mode determines how the Authorization Server returns result parameters from the Authorization Endpoint.
            public readonly string accessType = "offline";                                        // Indicates whether your application can refresh access tokens when the user is not present at the browser.
            public readonly string includeGrantedScopes = "true";                                 // Enables applications to use incremental authorization to request access to additional scopes in context.
            public readonly string state = "state_parameter_passthrough_value";                   // Specifies any string value that your application uses to maintain state between your authorization request and the authorization server's response.
            public readonly string grantType = "authorization_code";                              // “grant type” refers to the way an application gets an access token.
        }

        public OAuth2(string clientId, string clientSecret, string sucessPage, string unsucessPage, string pathToken)
        {
            oAuth2Info = new OAuth2Info();
            oAuth2Request = new OAuth2Request();
            oAuth2Profile = new OAuth2Profile();
            httpReq = new HttpReq();

            this.pathToken = pathToken;

            if (sucessPage != null)
            {
                oAuth2Info.sucessPage = sucessPage;
            }
            else
            {
                oAuth2Info.sucessPage = null;
            }

            if (unsucessPage != null)
            {
                oAuth2Info.unsucessPage = unsucessPage;
            }
            else
            {
                oAuth2Info.unsucessPage = null;
            }

            oAuth2Info.clientId = clientId;
            oAuth2Info.clientSecret = clientSecret;
        }
        public void StartTimer()
        {
            System.Timers.Timer t = new System.Timers.Timer(1800000);
            t.AutoReset = true;
            t.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            t.Start();
        }

        private async void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            await RefreshTokenAsync();
            Debug.WriteLine("TIMER:\n" + "refresh_token: " + oAuth2Info.refresh_token + "\naccess_token: " + oAuth2Info.access_token);
        }
        private async Task<bool> RefreshTokenAsync()
        {
            // Makes token request.
            string data =
                "&client_id=" + oAuth2Info.clientId +
                "&client_secret=" + oAuth2Info.clientSecret +
                "&refresh_token=" + oAuth2Info.refresh_token +
                "&grant_type=refresh_token" +
                "&access_type=offline" +
                "&prompt=consent";

            Dictionary<string, string> responseToken = await httpReq.SendPostRequestAsync("https://accounts.google.com/o/oauth2/token", data);

            if (responseToken.ContainsKey("error"))
            {
                return false;
            }
            else
            {
                oAuth2Info.access_token = responseToken["access_token"];
                oAuth2Info.expires_in = responseToken["expires_in"];
                File.WriteAllText(pathToken, JsonConvert.SerializeObject(oAuth2Info));

                return true;
            }
        }

        public OAuth2Profile GetProfile()
        {
            if (oAuth2Profile != null)
            {
                return oAuth2Profile;
            }
            else
            {
                return null;
            }
        }

        public async Task<bool> InitializeProfileAsync()
        {
            if (File.Exists(pathToken))
            {
                oAuth2Info = JsonConvert.DeserializeObject<OAuth2Info>(File.ReadAllText(pathToken));
                string data;
                try
                {
                    data = await httpReq.SendGetBearerAuthRequestAsync("https://gmail.googleapis.com/gmail/v1/users/me/profile", oAuth2Info.access_token);
                    Dictionary<string, string> profile = JsonConvert.DeserializeObject<Dictionary<string, string>>(data);

                    oAuth2Profile.EmailAddress = profile["emailAddress"];
                    oAuth2Profile.MessagesTotal = profile["messagesTotal"];
                    oAuth2Profile.ThreadsTotal = profile["threadsTotal"];
                    oAuth2Profile.HistoryId = profile["historyId"];

                    return true;
                }
                catch (WebException ex)
                {
                    Debug.WriteLine("WebExeption: " + ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public bool InitializeProfile()
        {
            if (File.Exists(pathToken))
            {
                oAuth2Info = JsonConvert.DeserializeObject<OAuth2Info>(File.ReadAllText(pathToken));
                string data;
                try
                {
                    data = httpReq.SendGetBearerAuthRequest("https://gmail.googleapis.com/gmail/v1/users/me/profile", oAuth2Info.access_token);
                    Dictionary<string, string> profile = JsonConvert.DeserializeObject<Dictionary<string, string>>(data);

                    oAuth2Profile.EmailAddress = profile["emailAddress"];
                    oAuth2Profile.MessagesTotal = profile["messagesTotal"];
                    oAuth2Profile.ThreadsTotal = profile["threadsTotal"];
                    oAuth2Profile.HistoryId = profile["historyId"];

                    return true;
                }
                catch (WebException ex)
                {
                    Debug.WriteLine("WebExeption: " + ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private bool RefreshToken()
        {
            // Makes token request.
            string data =
                "&client_id=" + oAuth2Info.clientId +
                "&client_secret=" + oAuth2Info.clientSecret +
                "&refresh_token=" + oAuth2Info.refresh_token +
                "&grant_type=refresh_token" +
                "&access_type=offline" +
                "&prompt=consent";

            Dictionary<string, string> responseToken = httpReq.SendPostRequest("https://accounts.google.com/o/oauth2/token", data);

            if (responseToken.ContainsKey("error"))
            {
                return false;
            }
            else
            {
                oAuth2Info.access_token = responseToken["access_token"];
                oAuth2Info.expires_in = responseToken["expires_in"];
                File.WriteAllText(pathToken, JsonConvert.SerializeObject(oAuth2Info));

                return true;
            }

        }
        public async Task<bool> CheckExistTokenAsync()
        {
            if (File.Exists(pathToken))
            {
                oAuth2Info = JsonConvert.DeserializeObject<OAuth2Info>(File.ReadAllText(pathToken));
                string data;
                try
                {
                    data = await httpReq.SendGetBearerAuthRequestAsync("https://gmail.googleapis.com/gmail/v1/users/me/profile", oAuth2Info.access_token);
                    return true;
                }
                catch (WebException ex)
                {
                    Debug.WriteLine("WebExeption: " + ex.Message);
                    if (await RefreshTokenAsync())
                    {
                        Debug.WriteLine("Refresh token was completed.");
                        return true;
                    }
                    else
                    {
                        Debug.WriteLine("Refresh token wasn\'t completed.");
                        return false;
                    }

                }
            }
            else
            {
                return false;
            }
        }
        public bool CheckExistToken()
        {
            if (File.Exists(pathToken))
            {
                //"https://gmail.googleapis.com/gmail/v1/users/me/profile"
                oAuth2Info = JsonConvert.DeserializeObject<OAuth2Info>(File.ReadAllText(pathToken));
                string data;
                try
                {
                    data = httpReq.SendGetBearerAuthRequest("https://gmail.googleapis.com/gmail/v1/users/me/profile", oAuth2Info.access_token);
                    return true;
                }
                catch (WebException ex)
                {
                    Debug.WriteLine("WebExeption: " + ex.Message);
                    if (RefreshToken())
                    {
                        Debug.WriteLine("Refresh token was completed.");
                        return true;
                    }
                    else
                    {
                        Debug.WriteLine("Refresh token wasn\'t completed.");
                        return false;
                    }

                }
            }
            else
            {
                return false;
            }
        }
        public async Task<bool> AuthenticationAsync()
        {

            Debug.WriteLine("Start authentication.");
            // Creates a redirect URI using an available port on the loopback address.
            oAuth2Info.redirectUri = $"http://{IPAddress.Loopback}:{GetRandomUnusedPort()}/";
            Debug.WriteLine("Redirect URI: " + oAuth2Info.redirectUri);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(oAuth2Info.redirectUri);
            Debug.WriteLine("Listening..");
            http.Start();
            // Creates the OAuth 2.0 authorization request.
            string authorizationRequest =
                oAuth2Request.requestAPI +
                "?response_type=" + oAuth2Request.responseType +
                "&scope=" + oAuth2Request.scope +
                "&access_type=" + oAuth2Request.accessType +
                "&include_granted_scopes=" + oAuth2Request.includeGrantedScopes +
                "&state=" + oAuth2Request.state +
                "&client_id=" + oAuth2Info.clientId +
                "&redirect_uri=" + oAuth2Info.redirectUri;

            // Opens request in the browser.
            Process.Start(authorizationRequest);

            // Waits for the OAuth authorization response.
            context = await http.GetContextAsync();

            // Makes token request.
            string data = "code=" + GetContextValue("code") + "&client_id=" + oAuth2Info.clientId + "&client_secret=" + oAuth2Info.clientSecret + "&redirect_uri=" + oAuth2Info.redirectUri + "&grant_type=" + oAuth2Request.grantType;

            Dictionary<string, string> tokenEndpointDecoded = await httpReq.SendPostRequestAsync("https://accounts.google.com/o/oauth2/token", data);
            if (tokenEndpointDecoded.ContainsKey("error"))
            {
                if (File.Exists(oAuth2Info.unsucessPage))
                {
                    Process.Start(oAuth2Info.unsucessPage);
                }
                return false;
            }
            else
            {
                oAuth2Info.access_token = tokenEndpointDecoded["access_token"];
                oAuth2Info.expires_in = tokenEndpointDecoded["expires_in"];
                oAuth2Info.refresh_token = tokenEndpointDecoded["refresh_token"];

                File.WriteAllText(pathToken, JsonConvert.SerializeObject(oAuth2Info));

                if (File.Exists(oAuth2Info.sucessPage))
                {
                    Process.Start(oAuth2Info.sucessPage);
                }
            }

            await RefreshTokenAsync();
            //StartTimer();
            return true;
        }
        public bool Authentication()
        {

            Debug.WriteLine("Start authentication.");
            // Creates a redirect URI using an available port on the loopback address.
            oAuth2Info.redirectUri = $"http://{IPAddress.Loopback}:{GetRandomUnusedPort()}/";
            Debug.WriteLine("Redirect URI: " + oAuth2Info.redirectUri);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(oAuth2Info.redirectUri);
            Debug.WriteLine("Listening..");
            http.Start();
            // Creates the OAuth 2.0 authorization request.
            string authorizationRequest =
                oAuth2Request.requestAPI +
                "?response_type=" + oAuth2Request.responseType +
                "&scope=" + oAuth2Request.scope +
                "&access_type=" + oAuth2Request.accessType +
                "&include_granted_scopes=" + oAuth2Request.includeGrantedScopes +
                "&state=" + oAuth2Request.state +
                "&client_id=" + oAuth2Info.clientId +
                "&redirect_uri=" + oAuth2Info.redirectUri;

            // Opens request in the browser.
            Process.Start(authorizationRequest);

            // Waits for the OAuth authorization response.
            context = http.GetContext();

            // Makes token request.
            string data = "code=" + GetContextValue("code") + "&client_id=" + oAuth2Info.clientId + "&client_secret=" + oAuth2Info.clientSecret + "&redirect_uri=" + oAuth2Info.redirectUri + "&grant_type=" + oAuth2Request.grantType;

            Dictionary<string, string> tokenEndpointDecoded = httpReq.SendPostRequest("https://accounts.google.com/o/oauth2/token", data);
            if (tokenEndpointDecoded.ContainsKey("error"))
            {
                if (File.Exists(oAuth2Info.unsucessPage))
                {
                    Process.Start(oAuth2Info.unsucessPage);
                }
                return false;
            }
            else
            {
                oAuth2Info.access_token = tokenEndpointDecoded["access_token"];
                oAuth2Info.expires_in = tokenEndpointDecoded["expires_in"];
                oAuth2Info.refresh_token = tokenEndpointDecoded["refresh_token"];

                File.WriteAllText(pathToken, JsonConvert.SerializeObject(oAuth2Info));

                if (File.Exists(oAuth2Info.sucessPage))
                {
                    Process.Start(oAuth2Info.sucessPage);
                }
            }

            RefreshToken();
            //StartTimer();
            return true;
        }
        public int Send(MimeMessage message)
        {
            var url = "https://www.googleapis.com/gmail/v1/users/me/messages/send";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";

            httpRequest.Accept = "application/json";
            httpRequest.ContentType = "application/json";
            httpRequest.Headers["Authorization"] = "Bearer " + oAuth2Info.access_token;
            //File.WriteAllText(pathToken, JsonConvert.SerializeObject(oAuth2Info));

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
    public class OAuth2Profile
    {
        private string emailAddress;
        private string messagesTotal;
        private string threadsTotal;
        private string historyId;

        public string EmailAddress
        {
            get
            {
                return emailAddress;
            }
            set
            {
                emailAddress = value;
            }
        }
        public string MessagesTotal
        {
            get
            {
                return messagesTotal;
            }
            set
            {
                messagesTotal = value;
            }
        }
        public string ThreadsTotal
        {
            get
            {
                return threadsTotal;
            }
            set
            {
                threadsTotal = value;
            }
        }
        public string HistoryId
        {
            get
            {
                return historyId;
            }
            set
            {
                historyId = value;
            }
        }
    }
}
