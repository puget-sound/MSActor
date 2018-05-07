using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class NewMoveRequestModel
    {
        public string identity;
        public string targetdatabase;

        public NewMoveRequestModel(string identity, string targetdatabase)
        {
            this.identity = identity;
            this.targetdatabase = targetdatabase;
        }
    }
}