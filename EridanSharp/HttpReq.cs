using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{

    class HttpReq : IInternetReq
    {
        private string method;
        private string accept;
        private string contentType;
        private string contentLength;

        public string SendGetBearerAuthRequest(string url, string token)
        {
            string data = "";
            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "GET";

            httpRequest.Accept = "application/json";
            httpRequest.ContentType = "application/json";
            httpRequest.Headers["Authorization"] = "Bearer " + token;

            using (WebResponse webResponse = httpRequest.GetResponse())
            using (Stream webStream = webResponse.GetResponseStream())
            using (StreamReader sr = new StreamReader(webStream))
            {
                data = sr.ReadToEnd();
            }

            return data;
        }

        public async Task<string> SendGetBearerAuthRequestAsync(string url, string token)
        {
            string data = "";
            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "GET";

            httpRequest.Accept = "application/json";
            httpRequest.ContentType = "application/json";
            httpRequest.Headers["Authorization"] = "Bearer " + token;

            using (WebResponse webResponse = await httpRequest.GetResponseAsync())
            using (Stream webStream = webResponse.GetResponseStream())
            using (StreamReader sr = new StreamReader(webStream))
            {
                data = await sr.ReadToEndAsync();
            }

            return data;
        }

        public string SendPostRequest(string url, string data)
        {
            throw new NotImplementedException();
        }

        public async Task<Dictionary<string, string>> SendPostRequestAsync(string url, string data)
        {
            Dictionary<string, string> tokenEndpointDecoded = new Dictionary<string, string>();
            // Sends the request
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(url);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            byte[] tokenRequestBodyBytes = Encoding.ASCII.GetBytes(data);
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
                    tokenEndpointDecoded = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseText);



                    return tokenEndpointDecoded;
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
                            tokenEndpointDecoded = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseText);
                            Debug.WriteLine(responseText);
                        }
                    }
                }
                return tokenEndpointDecoded;
            }
        }

        string IInternetReq.SendGetBearerAuthRequestAsync(string url, string token)
        {
            throw new NotImplementedException();
        }

        string IInternetReq.SendPostRequestAsync(string url, string data)
        {
            throw new NotImplementedException();
        }
    }
}
