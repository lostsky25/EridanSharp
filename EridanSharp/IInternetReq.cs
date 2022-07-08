using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{
    interface IInternetReq
    {
        string SendGetBearerAuthRequestAsync(string url, string token);
        string SendGetBearerAuthRequest(string url, string token);
        string SendPostRequestAsync(string url, string data);
        string SendPostRequest(string url, string data);
    }
}
