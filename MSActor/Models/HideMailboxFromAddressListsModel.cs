using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class HideMailboxFromAddressListsModel
    {
        public string identity;
        public string hidemailbox;

        public HideMailboxFromAddressListsModel(string identity, string hidemailbox)
        {
            this.identity = identity;
            this.hidemailbox = hidemailbox;
        }
    }
}