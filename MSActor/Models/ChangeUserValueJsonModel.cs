using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ChangeUserValueJsonModel
    {
        public string emplid;
        public string value;
        
        public ChangeUserValueJsonModel(string emplid, string value)
        {
            this.emplid = emplid;
            this.value = value;
        }
    }
}