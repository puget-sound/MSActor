using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class AddADGroupMemberModel
    {
        public string identity;
        public string member;

        public AddADGroupMemberModel(string group_identity, string group_member)
        {
            this.identity = group_identity;
            this.member = group_member;
        }
    }
}