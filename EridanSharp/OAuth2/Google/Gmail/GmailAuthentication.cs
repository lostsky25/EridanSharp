using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp.OAuth2.Google.Gmail
{
    class GmailAuthentication
    {
        public string access_token;             // Access token is the thing that applications use to make API requests on behalf of a user.
        public string refresh_token;            // Needs for updating current token.
        public string expires_in;               // How long to live\available current token.
        public string scope;                    // Scope is a mechanism in OAuth 2.0 to limit an application's access to a user's account.
        public string token_type;               // Type of the token.

        public string error;                    // 
        public string error_description;        // 
    }
}
