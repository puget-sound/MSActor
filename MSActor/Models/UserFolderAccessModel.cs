using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class UserFolderAccessModel
    {
        public string employeeid;
        public string samaccountname;
        public string computername;
        public string path;
        public string accesstype;

        public UserFolderAccessModel(string employeeid, string samaccountname, string computername, string path, string accesstype)
        {
            this.employeeid = employeeid;
            this.samaccountname = samaccountname;
            this.computername = computername;
            this.path = path;
            this.accesstype = accesstype;
        }
    }
}