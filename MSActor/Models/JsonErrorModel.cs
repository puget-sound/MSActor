using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    /// <summary>
    /// this is an object to handle returning an error as a json object
    /// to be used with the JsonConvert.SeralizeObject method
    /// </summary>
    public class JsonErrorModel
    {
        public string error;
        public JsonErrorModel(string error)
        {
            this.error = error;
        }
    }
}