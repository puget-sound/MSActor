using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class SetMailboxNameModel
    {
        public string identity;
        public string alias;
        public string addemailaddress;

        public SetMailboxNameModel(string identity, string alias, string addemailaddress)
        {
            this.identity = identity;
            this.alias = alias;
            this.addemailaddress = addemailaddress;
        }
    }
}