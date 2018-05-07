using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class IdentityModel
    {
        public string identity;

        public IdentityModel(string identity)
        {
            this.identity = identity;
        }
    }
}