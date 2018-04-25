using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ChangeUserValueModel
    {
        public string employeeid;
        public string samaccountname;
        public string field;
        public string value;
        
        public ChangeUserValueModel(string employeeid, string samaccountname, string field, string value)
        {
            
            this.employeeid = employeeid;
            this.samaccountname = samaccountname;
            this.field = field;
            this.value = value;

        }
    }
}