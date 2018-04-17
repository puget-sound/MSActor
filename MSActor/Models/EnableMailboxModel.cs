using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class EnableMailboxModel
    {
        public string emailaddresses;
        public string database;
        public string alias;
        public EnableMailboxModel(string database, string alias, string emailaddresses)
        {
            this.emailaddresses = emailaddresses;
            this.database = database;
            this.alias = alias;
        }
    }
}