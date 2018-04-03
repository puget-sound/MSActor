using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using MSActor.Models;
using Newtonsoft.Json;


namespace MSActor.Controllers
{
    public class MSActorController : ApiController
    { 
        /// <summary>
        /// Constructor for the MSActor Controller. Runs all MSActor Functions. 
        /// </summary>
        public MSActorController() {

        }

        /// <summary>
        /// This is a method to return the information about a user based on emplid. 
        /// Parameter is a Json object with the form {emplid = XXXXXXXXX}
        /// Only used for testing, not intended to be used in production. 
        /// </summary>
        [Route("getaduserbyemplid")]
        [HttpPost]
        public ADUserModel GetAdUserByEmplid([FromBody] EmplidModel emplidWrap)
        {
            ADController control = new ADController();
            return control.GetADUserDriver(emplidWrap.emplid);
        }

        /// <summary>$
        /// This is a method for creating new users in AD. Calls a driver method in the ADController. 
        /// </summary>
        [Route("newaduser")]
        [HttpPost]
        public MSActorReturnMessageModel NewADUser([FromBody] ADUserModel newUser)
        {
            ADController control = new ADController();
            return control.NewADUserDriver(newUser);
        }


        /// <summary>
        /// This method changes the surname of a user in AD. 
        /// </summary>
        /// <param name="input"> This is a json parameter with the form 
        /// 
        /// {
        ///     emplid = "XXXXXXXXXX",
        ///     value  = "Value Here"
        /// }
        /// </param>
        /// <returns></returns>
        [Route("changeusersurname")]
        [HttpPost]
        public MSActorReturnMessageModel ChangeUserSurname([FromBody] ChangeUserValueJsonModel input)
        {
            ADController control = new ADController();
            return control.ChangeUserSurnameDriver(input.emplid, input.value);
        }
    }
}
