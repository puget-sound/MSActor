using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class SetMailboxQuotasModel
    {
        public string identity;
        public string prohibitsendreceivequota;
        public string prohibitsendquota;
        public string issuewarningquota;

        public SetMailboxQuotasModel(string identity, string prohibitsendreceivequota, string prohibitsendquota, string issuewarningquota)
        {
            this.identity = identity;
            this.prohibitsendreceivequota = prohibitsendreceivequota;
            this.prohibitsendquota = prohibitsendquota;
            this.issuewarningquota = issuewarningquota;
        }
    }
}