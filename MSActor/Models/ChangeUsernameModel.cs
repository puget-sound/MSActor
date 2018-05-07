using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ChangeUsernameModel
    {
        public string employeeid;
        public string searchbase;
        public string samaccountname;
        public string userprincipalname;

        public ChangeUsernameModel(string employeeid, string searchbase, string samaccountname, string userprincipalname)
        {
            this.employeeid = employeeid;
            this.searchbase = searchbase;
            this.samaccountname = samaccountname;
            this.userprincipalname = userprincipalname;
        }
    }
}