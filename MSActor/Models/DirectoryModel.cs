using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class DirectoryModel
    {
        public string computername;
        public string path;

        public DirectoryModel(string computername, string path)
        {
            this.computername = computername;
            this.path = path;
        }
    }
}