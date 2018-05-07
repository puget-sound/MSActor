using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class RemoveADGroupModel
    {
        public string identity;

        public RemoveADGroupModel(string group_identity)
        {
            this.identity = group_identity;
        }
    }
}