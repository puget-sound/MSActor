using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    /// <summary>
    /// This class holds the code and message to return to PeopleSoft from the MSActor
    /// </summary>
    public class MSActorReturnMessageModel
    {

        public string code; 
        public string message;
        public MSActorReturnMessageModel(string code, string message)
        {
            this.code = code;
            this.message = message;
        }
    }
}