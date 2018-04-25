using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class CreateFolderModel
    {
        public string path;
        public CreateFolderModel(string path)
        {
            this.path = path;
        }
    }
}