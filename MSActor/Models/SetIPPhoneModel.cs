using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class SetIPPhoneModel
    {
        public string employeeid;
        public string samaccountname;
        public string ipphone;

        public SetIPPhoneModel(string employeeid, string samaccountname, string ipphone)
        {
            this.employeeid = employeeid;
            this.samaccountname = samaccountname;
            this.ipphone = ipphone;
        }
    }
}