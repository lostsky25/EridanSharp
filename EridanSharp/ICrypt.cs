using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{
    internal interface ICrypt
    {
        string Base64Encode(string data);
        string Base64Decode(string data);
    }
}
