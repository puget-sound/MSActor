using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ChangeUserValueJsonModel
    {
        public string emplid;
        public string field;
        public string value;
        
        public ChangeUserValueJsonModel(string emplid, string field, string value)
        {
            this.emplid = emplid;
            this.field = field;
            this.value = value;
        }
    }
}