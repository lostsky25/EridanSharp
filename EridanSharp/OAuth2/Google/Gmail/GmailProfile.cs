using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{
    public class GmailProfile
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
