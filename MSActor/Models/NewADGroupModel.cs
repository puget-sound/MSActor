using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class NewADGroupModel
    {
        public string name;
        public string description;
        public string info;
        public string path;
        public string groupcategory;
        public string groupscope;

        public NewADGroupModel(string group_name, string group_description, string group_info, string group_ad_path, string group_category, string group_scope)
        {
            this.name = group_name;
            this.description = group_description;
            this.info = group_info;
            this.path = group_ad_path;
            this.groupcategory = group_category;
            this.groupscope = group_scope;
        }
    }
}