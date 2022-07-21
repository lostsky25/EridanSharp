using EridanSharp.OAuth2.Google.Gmail;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace EridanSharp
{
    namespace OAuth2 { 
        public class Gmail
        {
            #region mainRequests
                private const string _req_profile       = "https://gmail.googleapis.com/gmail/v1/users/me/profile";
                private const string _req_auth          = "https://accounts.google.com/o/oauth2/v2/auth";
                private const string _req_scope         = "https://www.googleapis.com/auth/gmail.modify";
                private const string _req_token         = "https://accounts.google.com/o/oauth2/token";
                private const string _req_send          = "https://www.googleapis.com/gmail/v1/users/me/messages/send";
            #endregion

            private GmailProfile gmailProfile;
            private GmailAuthentication gmailAuthentication;

            private HttpListenerContext context;
            private GmailRequest oAuth2Request;
            private HttpReq httpReq;
            private string pathToken;
            private System.Timers.Timer timerRefresh;

            #region timerConfiguration
                private bool timerAutoRefresh = true;
                /// <summary>
                /// The default value is true. Your token will
                /// refresh every 30 min. You can disable this 
                /// option if you wish just set it as false.
                /// </summary>
                public bool TimerAutoRefresh
                {
                    get
                    {
                        return timerAutoRefresh;
                    }
                    set
                    {
                        timerAutoRefresh = value;
                    }
                }

                private uint timerRefreshInterval = 1800000;
                /// <summary>
                /// Set the value as ms (1000ms is equal 1sec). 
                /// The default value is 30 min (1800000ms). It means 
                /// that your token will refresh every 30 min or every 
                /// 1800000ms If you turned on auto refresh before.
                /// </summary>
                public uint TimerRefreshInterval
                {
                    get
                    {
                        return timerRefreshInterval;
                    }
                    set
                    {
                        if (timerAutoRefresh)
                        {
                            timerRefreshInterval = value;
                        }
                        else
                        {
                            throw new InvalidOperationException("TimerAutoRefresh is false. You can't use TimerRefreshInterval.");
                        }
                    }
                }
            #endregion

            private string clientId;
            private string clientSecret;
            private string redirectUri;
            private string sucessPage;
            private string unsucessPage;

            private class GmailRequest
            {
                public readonly string requestAPI = _req_auth;                                  // The beginning part of the authorization request.
                public readonly string scope = _req_scope;                                      // Scope is a mechanism in OAuth 2.0 to limit an application's access to a user's account.
                public readonly string responseType = "code";                                   // The Response Mode determines how the Authorization Server returns result parameters from the Authorization Endpoint.
                public readonly string accessType = "offline";                                  // Indicates whether your application can refresh access tokens when the user is not present at the browser.
                public readonly string includeGrantedScopes = "true";                           // Enables applications to use incremental authorization to request access to additional scopes in context.
                public readonly string state = "state_parameter_passthrough_value";             // Specifies any string value that your application uses to maintain state between your authorization request and the authorization server's response.
                public readonly string grantType = "authorization_code";                        // “grant type” refers to the way an application gets an access token.
            }

            public Gmail(string clientId, string clientSecret, string sucessPage, string unsucessPage, string pathToken)
            {
                gmailAuthentication = new GmailAuthentication();

                oAuth2Request = new GmailRequest();
                gmailProfile = new GmailProfile();
                httpReq = new HttpReq();

                this.pathToken = pathToken;

                if (sucessPage != null)
                {
                    this.sucessPage = sucessPage;
                }

                if (unsucessPage != null)
                {
                    this.unsucessPage = unsucessPage;
                }

                this.clientId = clientId;
                this.clientSecret = clientSecret;
            }

            #region NonAsync
                public bool Authentication()
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
                    string authorizationRequest =
                        oAuth2Request.requestAPI +
                        "?response_type=" + oAuth2Request.responseType +
                        "&scope=" + oAuth2Request.scope +
                        "&access_type=" + oAuth2Request.accessType +
                        "&include_granted_scopes=" + oAuth2Request.includeGrantedScopes +
                        "&state=" + oAuth2Request.state +
                        "&client_id=" + clientId +
                        "&redirect_uri=" + redirectUri;

                    // Opens request in the browser.
                    Process.Start(authorizationRequest);

                    // Waits for the OAuth authorization response.
                    context = http.GetContext();

                    // Makes token request.
                    string data =
                        "code=" + GetContextValue("code") +
                        "&client_id=" + clientId +
                        "&client_secret=" + clientSecret +
                        "&redirect_uri=" + redirectUri +
                        "&grant_type=" + oAuth2Request.grantType;

                    var response = JsonConvert.DeserializeObject<Dictionary<string, string>>(httpReq.SendPostRequest(_req_token, data));
                    if (response.ContainsKey("error"))
                    {
                        return false;
                    }
                    else
                    {
                        gmailAuthentication.access_token = response["access_token"];
                        gmailAuthentication.refresh_token = response["refresh_token"];
                        gmailAuthentication.expires_in = response["expires_in"];
                        gmailAuthentication.scope = response["scope"];
                        gmailAuthentication.token_type = response["token_type"];

                        File.WriteAllText(pathToken, JsonConvert.SerializeObject(gmailAuthentication));
                    }

                    gmailProfile = JsonConvert.DeserializeObject<GmailProfile>(httpReq.SendGetBearerAuthRequest(_req_profile, gmailAuthentication.access_token));
                    RefreshToken();
                    if (timerRefresh == null)
                    {
                        StartTimer();
                    }
                    return true;
                }
                public bool CheckExistToken()
                {
                    if (File.Exists(pathToken))
                    {
                        gmailAuthentication = JsonConvert.DeserializeObject<GmailAuthentication>(File.ReadAllText(pathToken));
                        try
                        {
                            gmailProfile = JsonConvert.DeserializeObject<GmailProfile>(httpReq.SendGetBearerAuthRequest(_req_profile, gmailAuthentication.access_token));
                            if (timerRefresh == null)
                            {
                                StartTimer();
                            }
                            return true;
                        }
                        catch (WebException ex)
                        {
                            Debug.WriteLine("WebExeption: " + ex.Message);
                            if (RefreshToken())
                            {
                                if (timerRefresh == null)
                                {
                                    StartTimer();
                                }
                                return true;
                            }
                            else
                            {
                                return false;
                            }
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
                        try
                        {
                            gmailProfile = JsonConvert.DeserializeObject<GmailProfile>(httpReq.SendGetBearerAuthRequest(_req_profile, gmailAuthentication.access_token));

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
                        "&client_id=" + clientId +
                        "&client_secret=" + clientSecret +
                        "&refresh_token=" + gmailAuthentication.refresh_token +
                        "&grant_type=refresh_token" +
                        "&access_type=offline" +
                        "&prompt=consent";

                    var response = JsonConvert.DeserializeObject<Dictionary<string, string>>(httpReq.SendPostRequest(_req_token, data));

                    if (response.ContainsKey("error"))
                    {
                        return false;
                    }
                    else
                    {
                        gmailAuthentication.access_token = response["access_token"];
                        gmailAuthentication.expires_in = response["expires_in"];
                        gmailAuthentication.scope = response["scope"];
                        gmailAuthentication.token_type = response["token_type"];
                        File.WriteAllText(pathToken, JsonConvert.SerializeObject(gmailAuthentication));
                        return true;
                    }
                }
                public void Send(MimeMessage message)
                {
                    var url = _req_send;

                    var httpRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpRequest.Method = "POST";

                    httpRequest.Accept = "application/json";
                    httpRequest.ContentType = "application/json";
                    httpRequest.Headers["Authorization"] = "Bearer " + gmailAuthentication.access_token;

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

                    return;
                }
            #endregion

            #region Async
            public async Task<bool> AuthenticationAsync()
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
                        string authorizationRequest =
                            oAuth2Request.requestAPI +
                            "?response_type=" + oAuth2Request.responseType +
                            "&scope=" + oAuth2Request.scope +
                            "&access_type=" + oAuth2Request.accessType +
                            "&include_granted_scopes=" + oAuth2Request.includeGrantedScopes +
                            "&state=" + oAuth2Request.state +
                            "&client_id=" + clientId +
                            "&redirect_uri=" + redirectUri;

                        // Opens request in the browser.
                        Process.Start(authorizationRequest);

                        // Waits for the OAuth authorization response.
                        context = await http.GetContextAsync();

                    // Makes token request.
                    string data =
                        "code=" + GetContextValue("code") +
                        "&client_id=" + clientId +
                        "&client_secret=" + clientSecret +
                        "&redirect_uri=" + redirectUri +
                        "&grant_type=" + oAuth2Request.grantType;

                    var response = JsonConvert.DeserializeObject<Dictionary<string, string>>(await httpReq.SendPostRequestAsync(_req_token, data));
                    if (response.ContainsKey("error"))
                    {
                        return false;
                    }
                    else
                    {
                        gmailAuthentication.access_token = response["access_token"];
                        gmailAuthentication.refresh_token = response["refresh_token"];
                        gmailAuthentication.expires_in = response["expires_in"];
                        gmailAuthentication.scope = response["scope"];
                        gmailAuthentication.token_type = response["token_type"];

                        File.WriteAllText(pathToken, JsonConvert.SerializeObject(gmailAuthentication));
                    }

                    gmailProfile = JsonConvert.DeserializeObject<GmailProfile>(await httpReq.SendGetBearerAuthRequestAsync(_req_profile, gmailAuthentication.access_token));
                    await RefreshTokenAsync();
                    if (timerRefresh == null)
                    {
                        StartTimer();
                    }
                    return true;
                }
                public async Task<bool> CheckExistTokenAsync()
                {
                    if (File.Exists(pathToken))
                    {
                        gmailAuthentication = JsonConvert.DeserializeObject<GmailAuthentication>(File.ReadAllText(pathToken));
                        try
                        {
                            gmailProfile = JsonConvert.DeserializeObject<GmailProfile>(httpReq.SendGetBearerAuthRequest(_req_profile, gmailAuthentication.access_token));
                            if (timerRefresh == null)
                            {
                                StartTimer();
                            }
                            return true;
                        }
                        catch (WebException ex)
                        {
                            Debug.WriteLine("WebExeption: " + ex.Message);
                            if (await RefreshTokenAsync())
                            {
                                Debug.WriteLine("Refresh token was completed.");
                                if (timerRefresh == null)
                                {
                                    StartTimer();
                                }
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
                public async Task<bool> InitializeProfileAsync()
                {
                    if (File.Exists(pathToken))
                    {
                        try
                        {
                            gmailProfile = JsonConvert.DeserializeObject<GmailProfile>(await httpReq.SendGetBearerAuthRequestAsync(_req_profile, gmailAuthentication.access_token));

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
                private async Task<bool> RefreshTokenAsync()
                {
                    // Makes token request.
                    string data =
                        "&client_id=" + clientId +
                        "&client_secret=" + clientSecret +
                        "&refresh_token=" + gmailAuthentication.refresh_token +
                        "&grant_type=refresh_token" +
                        "&access_type=offline" +
                        "&prompt=consent";

                    var response = JsonConvert.DeserializeObject<Dictionary<string, string>>(await httpReq.SendPostRequestAsync(_req_token, data));

                    if (response.ContainsKey("error"))
                    {
                        return false;
                    }
                    else
                    {
                        gmailAuthentication.access_token = response["access_token"];
                        gmailAuthentication.expires_in = response["expires_in"];
                        gmailAuthentication.scope = response["scope"];
                        gmailAuthentication.token_type = response["token_type"];
                        File.WriteAllText(pathToken, JsonConvert.SerializeObject(gmailAuthentication));
                        return true;
                    }
                }
            #endregion

            public void StartTimer()
            {
                if (timerAutoRefresh)
                {
                    timerRefresh = new System.Timers.Timer(timerRefreshInterval);
                    timerRefresh.AutoReset = true;
                    timerRefresh.Elapsed += new ElapsedEventHandler(OnTimedEvent);
                    RefreshToken();
                    timerRefresh.Start();
                }
            }

            private async void OnTimedEvent(Object source, ElapsedEventArgs e)
            {
                await RefreshTokenAsync();
                Debug.WriteLine("TIMER:\n" + "refresh_token: " + gmailAuthentication.refresh_token + "\naccess_token: " + gmailAuthentication.access_token);
            }
            
            public GmailProfile GetProfile()
            {
                if (gmailProfile != null)
                {
                    return gmailProfile;
                }
                else
                {
                    return null;
                }
            }

            public void ShowSucessPage()
            {
                if (File.Exists(sucessPage))
                {
                    Process.Start(sucessPage);
                }
            }

            public void ShowUnsucessPage()
            {
                if (File.Exists(unsucessPage))
                {
                    Process.Start(unsucessPage);
                }
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
}
