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
        UtilityController util;
        public MSActorController() {
            util = new UtilityController(); 
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

        /// <summary>
        /// This is a method for creating new users in AD. Calls a driver method in the ADController. 
        /// </summary>
        [Route("newaduser")]
        [HttpPost]
        public MSActorReturnMessageModel NewADUser([FromBody] ADUserModel newUser)
        {
            try { 

                ADController control = new ADController();
                return control.NewADUserDriver(newUser);
            }catch(Exception e)
            {
               return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.ChangeUserValueDriver(input.employeeid, input.samaccountname, input.field, input.value);
            }catch(Exception e)
            {
               return util.ReportError(e);
            }
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
            try { 
                ExchangeController control = new ExchangeController();
                return control.EnableMailbox(input.database, input.alias, input.emailaddresses);
            }catch(Exception e)
            {
               return util.ReportError(e);
            }
        }

        /// <summary>
        /// Creates a Folder on the file server at the given file path
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("newdirectory")]
        [HttpPost]
        public MSActorReturnMessageModel NewDirectory([FromBody] DirectoryModel input)
        {
            try
            {
                FileServerController control = new FileServerController();
                return control.NewDirectory(input.computername, input.path);
            }catch(Exception e)
            {
               return util.ReportError(e);
            }
        }

        /// <summary>
        /// Delete folder specified
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("removedirectory")]
        [HttpPost]
        public MSActorReturnMessageModel RemoveDirectory([FromBody] DirectoryModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.RemoveDirectory(input.computername, input.path);
            }catch(Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Rename folder specified
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("renamedirectory")]
        [HttpPost]
        public MSActorReturnMessageModel RenameDirectory([FromBody] RenameDirectoryModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.RenameDirectory(input.computername, input.path, input.newname);
            }catch(Exception e)
            {
               return util.ReportError(e);
            }
        }

        /// <summary>
        /// Create share for folder specified
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("addnetshare")]
        [HttpPost]
        public MSActorReturnMessageModel AddNetShare([FromBody] ShareModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.AddNetShare(input.name, input.computername, input.path);
            }catch(Exception e)
            {
               return util.ReportError(e);
            }
        }

        /// <summary>
        /// Delete share for folder specified
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("removenetshare")]
        [HttpPost]
        public MSActorReturnMessageModel RemoveNetShare([FromBody] ShareModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.RemoveNetShare(input.name, input.computername, input.path);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Grant specified access to this folder for specified User
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("adduserfolderaccess")]
        [HttpPost]
        public MSActorReturnMessageModel AddUserFolderAccess([FromBody] UserFolderAccessModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.AddUserFolderAccess(input.employeeid, input.samaccountname, input.computername, input.path, input.accesstype);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Grant specified access to this folder for specified group
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("addgroupfolderaccess")]
        [HttpPost]
        public MSActorReturnMessageModel AddGroupFolderAccess([FromBody] GroupFolderAccessModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.AddGroupFolderAccess(input.groupname, input.computername, input.path, input.accesstype);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Add quota on specified folder
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("adddirquota")]
        [HttpPost]
        public MSActorReturnMessageModel AddDirQuota([FromBody] DirectoryQuotaModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.AddDirQuota(input.computername, input.path, input.limit);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Modify quota on specified folder
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("modifydirquota")]
        [HttpPost]
        public MSActorReturnMessageModel ModifyDirQuota([FromBody] DirectoryQuotaModel input)
        {
            try { 
                FileServerController control = new FileServerController();
                return control.ModifyDirQuota(input.computername, input.path, input.limit);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.SetHomeDirectory(input.employeeid, input.samaccountname, input.homedirectory, input.homedrive);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.NewADGroup(input.name, input.description, input.info, input.path, 
                    input.groupcategory, input.groupscope, input.samaccountname);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.RemoveADGroup(input.identity);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.AddADGroupMember(input.identity, input.member);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.RemoveADGroupMember(input.identity, input.member);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.SetPassword(input.employeeid, input.samaccountname, input.accountpassword, input.changepasswordatlogon);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.RemoveADObject(input.employeeid, input.samaccountname);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.ChangeUsername(input.employeeid, input.old_samaccountname, input.new_samaccountname, input.userprincipalname);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
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
            try { 
                ADController control = new ADController();
                return control.SetIPPhone(input.employeeid, input.samaccountname, input.ipphone);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Sets quotas on mailboxes
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("setmailboxquotas")]
        [HttpPost]
        public MSActorReturnMessageModel SetMailboxQuotas([FromBody] SetMailboxQuotasModel input)
        {
            try { 
                ExchangeController control = new ExchangeController();
                return control.SetMailboxQuotas(input.identity, input.prohibitsendreceivequota, input.prohibitsendquota, input.issuewarningquota);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Changes alias on mailbox
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("setmailboxname")]
        [HttpPost]
        public MSActorReturnMessageModel SetMailboxName([FromBody] SetMailboxNameModel input)
        {
            try { 
                ExchangeController control = new ExchangeController();
                return control.SetMailboxName(input.identity, input.alias, input.addemailaddress);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Makes a new request to move mailbox from one database to another
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("newmoverequest")]
        [HttpPost]
        public MSActorReturnMessageModel NewMoveRequest([FromBody] NewMoveRequestModel input)
        {
            try { 
                ExchangeController control = new ExchangeController();
                return control.NewMoveRequest(input.identity, input.targetdatabase);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Checks on status of mailbox move request
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("getmoverequest")]
        [HttpPost]
        public MSActorReturnMessageModel GetMoveRequest([FromBody] IdentityModel input)
        {
            try { 
                ExchangeController control = new ExchangeController();
                return control.GetMoveRequest(input.identity);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Removes a mailbox
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("disablemailbox")]
        [HttpPost]
        public MSActorReturnMessageModel DisableMailbox([FromBody] IdentityModel input)
        {
            try { 
                ExchangeController control = new ExchangeController();
                return control.DisableMailbox(input.identity);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Sets flag on a mailbox to hide it from address lists
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [Route("hidemailboxfromaddresslists")]
        [HttpPost]
        public MSActorReturnMessageModel HideMailboxFromAddressLists([FromBody] HideMailboxFromAddressListsModel input)
        {
            try { 
                ExchangeController control = new ExchangeController();
                return control.HideMailboxFromAddressLists(input.identity, input.hidemailbox);
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /*[Route("enabledistributiongroup")]
        [HttpPost]
        public MSActorReturnMessageModel EnableDistributionGroup([FromBody] EnableDistributionGroupModel input)
        {
            try
            {
                ExchangeController control = new ExchangeController();
                return control.EnableDistributionGroup(input.identity);
            }
            catch(Exception e)
            {
                return util.ReportError(e);
            }
        }*/
    }
}
