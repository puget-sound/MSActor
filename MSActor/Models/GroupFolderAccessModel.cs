using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class GroupFolderAccessModel
    {
        public string groupname;
        public string computername;
        public string path;
        public string accesstype;

        public GroupFolderAccessModel(string groupname, string computername, string path, string accesstype)
        {
            this.groupname = groupname;
            this.computername = computername;
            this.path = path;
            this.accesstype = accesstype;
        }
    }
}