using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class RenameDirectoryModel
    {
        public string computername;
        public string path;
        public string newname;

        public RenameDirectoryModel(string computername, string path, string newname)
        {
            this.computername = computername;
            this.path = path;
            this.newname = newname;
        }
    }
}