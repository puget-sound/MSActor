using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class DirectoryQuotaModel
    {
        public string computername;
        public string path;
        public string limit;

        public DirectoryQuotaModel(string computername, string path, string limit)
        {
            this.computername = computername;
            this.path = path;
            this.limit = limit;
        }
    }
}