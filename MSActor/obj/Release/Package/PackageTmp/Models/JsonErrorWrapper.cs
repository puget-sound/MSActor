using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ContactManager.Models
{
    /// <summary>
    /// this is an object to handle returning an error as a json object
    /// to be used with the JsonConvert.SeralizeObject method
    /// </summary>
    public class JsonErrorWrapper
    {
        public string error;
        public JsonErrorWrapper(string error)
        {
            this.error = error;
        }
    }
}