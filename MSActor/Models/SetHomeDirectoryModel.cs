using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class SetHomeDirectoryModel
    {
        public string employeeid;
        public string samaccountname;
        public string homedirectory;
        public string homedrive;

        public SetHomeDirectoryModel(string employeeid, string samaccountname, string homedirectory, string homedrive)
        {
            this.employeeid = employeeid;
            this.samaccountname = samaccountname;
            this.homedirectory = homedirectory;
            this.homedrive = homedrive;
        }
    }
}