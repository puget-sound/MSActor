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
        public string old_samaccountname;
        public string new_samaccountname;
        public string userprincipalname;

        public ChangeUsernameModel(string employeeid, string searchbase, string old_samaccountname, string new_samaccountname, string userprincipalname)
        {
            this.employeeid = employeeid;
            this.searchbase = searchbase;
            this.old_samaccountname = old_samaccountname;
            this.new_samaccountname = new_samaccountname;
            this.userprincipalname = userprincipalname;
        }
    }
}