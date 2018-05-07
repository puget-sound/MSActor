using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ShareModel
    {
        public string name;
        public string computername;
        public string path;

        public ShareModel(string name, string computername, string path)
        {
            this.name = name;
            this.computername = computername;
            this.path = path;
        }
    }
}