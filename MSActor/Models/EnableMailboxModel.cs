using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class EnableMailboxModel
    {
        public string identity;
        public string database;
        public string alias;
        public EnableMailboxModel(string identity, string database, string alias)
        {
            this.identity = identity;
            this.database = database;
            this.alias = alias;
        }
    }
}