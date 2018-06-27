using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class EnableDistributionGroupModel
    {

        public string identity;

        public EnableDistributionGroupModel(string identity)
        {
            this.identity = identity;
        }
    }
}