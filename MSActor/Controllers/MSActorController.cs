﻿using System;
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
            Debug.WriteLine("We are Legion for we are many");
        }

        /// <summary>
        /// This is a method to return the information about a user based on emplid. 
        /// Parameter is a Json object with the form {emplid = XXXXXXXXX}
        /// Only used for testing, not intended to be used in production. 
        /// </summary>
        [Route("getaduserbyemplid")]
        [HttpPost]
        private ADUserModel GetAdUserByEmplid([FromBody] EmplidModel emplidWrap)
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
        /// This method changes the value of a field on a user in AD. 
        /// </summary>
        /// <param name="input"> This is a json parameter with the form 
        /// 
        /// {
        ///     emplid = "XXXXXXXXXX",
        ///     field  = "field name",
        ///     value  = "Value Here"
        /// }
        /// </param>
        /// <returns></returns>
        [Route("changeuservalue")]
        [HttpPost]
        public MSActorReturnMessageModel ChangeUserValue([FromBody] ChangeUserValueModel input)
        {
            ADController control = new ADController();
            Debug.WriteLine("General Kenobi: " + input.employeeid);
            return control.ChangeUserValueDriver(input.employeeid, input.samaccountname, input.field, input.value);
        }

        /// <summary>
        /// Activates an exchange mailbox for a user based on 
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("enablemailbox")]
        [HttpPost]
        public MSActorReturnMessageModel EnableMailbox([FromBody] EnableMailboxModel input)
        {
            ExchangeController control = new ExchangeController();
            return control.EnableMailboxDriver(input.database, input.alias, input.emailaddresses);
        }

        /// <summary>
        /// Creates a Folder on the file server at the given file path
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("newdirectory")]
        [HttpPost]
        public MSActorReturnMessageModel NewDirectory([FromBody] CreateFolderModel input)
        {
            FileServerController control = new FileServerController();
            return control.NewDirectory(input.path);
        }

        /// <summary>
        /// Creates a new Home Directory for the user
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("sethomedirectory")]
        [HttpPost]
        public MSActorReturnMessageModel SetHomeDirectory([FromBody] SetHomeDirectoryModel input)
        {
            ADController control = new ADController();
            return control.SetHomeDirectory(input.employeeid, input.samaccountname, input.homedirectory, input.homedrive);
        }

        /// <summary>
        /// Creates a new AD group
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("newadgroup")]
        [HttpPost]
        public MSActorReturnMessageModel NewADGroup([FromBody] NewADGroupModel input)
        {
            ADController control = new ADController();
            return control.NewADGroup(input.name, input.description, input.info, input.path, input.groupcategory, input.groupscope);
        }

        /// <summary>
        /// Removes an existing AD group
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("removeadgroup")]
        [HttpPost]
        public MSActorReturnMessageModel RemoveADGroup([FromBody] RemoveADGroupModel input)
        {
            ADController control = new ADController();
            return control.RemoveADGroup(input.identity);
        }

        /// <summary>
        /// Add user to an AD group
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("addadgroupmember")]
        [HttpPost]
        public MSActorReturnMessageModel AddADGroupMember([FromBody] AddADGroupMemberModel input)
        {
            ADController control = new ADController();
            return control.AddADGroupMember(input.identity, input.member);
        }

        /// <summary>
        /// Remove user from an AD group
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("removeadgroupmember")]
        [HttpPost]
        public MSActorReturnMessageModel RemoveADGroupMember([FromBody] RemoveADGroupMemberModel input)
        {
            ADController control = new ADController();
            return control.RemoveADGroupMember(input.identity, input.member);
        }

        /// <summary>
        /// Set password
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("setpassword")]
        [HttpPost]
        public MSActorReturnMessageModel SetPassword([FromBody] SetPasswordModel input)
        {
            ADController control = new ADController();
            return control.SetPassword(input.employeeid, input.samaccountname, input.accountpassword, input.changepasswordatlogon);
        }

        /// <summary>
        /// ...
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("removeadobject")]
        [HttpPost]
        public MSActorReturnMessageModel RemoveADObject([FromBody] RemoveADObjectModel input)
        {
            ADController control = new ADController();
            return control.RemoveADObject(input.employeeid, input.samaccountname);
        }

        /// <summary>
        /// ...
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("changeusername")]
        [HttpPost]
        public MSActorReturnMessageModel ChangeUsername([FromBody] ChangeUsernameModel input)
        {
            ADController control = new ADController();
            return control.ChangeUsername(input.employeeid, input.searchbase, input.samaccountname, input.userprincipalname);
        }

        /// <summary>
        /// ...
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("setipphone")]
        [HttpPost]
        public MSActorReturnMessageModel SetIPPhone([FromBody] SetIPPhoneModel input)
        {
            ADController control = new ADController();
            return control.SetIPPhone(input.employeeid, input.samaccountname, input.ipphone);
        }
    }
}
