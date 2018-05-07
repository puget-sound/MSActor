using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class SetPasswordModel
    {
        public string employeeid;
        public string samaccountname;
        public string accountpassword;
        public string changepasswordatlogon;

        public SetPasswordModel(string employeeid, string samaccountname, string accountpassword, string changepasswordatlogon)
        {
            this.employeeid = employeeid;
            this.samaccountname = samaccountname;
            this.accountpassword = accountpassword;
            this.changepasswordatlogon = changepasswordatlogon;
        }
    }
}