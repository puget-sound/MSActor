using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class RemoveADObjectModel
    {
        public string employeeid;
        public string samaccountname;

        public RemoveADObjectModel(string employeeid, string samaccountname)
        {
            this.employeeid = employeeid;
            this.samaccountname = samaccountname;
        }
    }
}