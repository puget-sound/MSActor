using MSActor.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace MSActor.Controllers
{
    /// <summary>
    /// This class holds all the methods used to operate on Active Directory in the MSActor. 
    /// There are methods in this actor for changing every value of a user in AD individually. 
    /// We (Jeff Strong and Robert Shelton) made this decision because we decided that it would be easier 
    /// rather than passing the field you wanted to edit as a parameter. 
    /// Author - Robert Shelton
    /// Date - 4/2/2018
    /// </summary>
    public class ADController
    {
        UtilityController util;
        public ADController()
        {
            util = new UtilityController();
        }
        public const string SuccessCode = "CMP";
        public const string ErrorCode = "ERR";
        public ADUserModel GetADUserDriver(string emplid)
        {
            string searchName = "";
            string searchCity = "";
            string searchCountry = "";
            string searchDepartment = "";
            string searchDescription = "";
            string searchDisplayName = "";
            string searchEmployeeID = "";
            string searchGivenName = "";
            string searchOfficePhone = "";
            string searchInitials = "";
            string searchOffice = "";
            string searchSamAccountName = "";
            string searchState = "";
            string searchStreetAddress = "";
            string searchSurname = "";
            string searchTitle = "";
            string searchObjectClass = "";
            string searchUserPrincipalName = "";
            string searchPath = "";
            string searchPostalCode = "";
            string searchType = "";
            string searchIPPhone = "";
            string searchMSExchHideFromAddressList = "";
            string searchChangePasswordAtLogon = "";

            string searchEnabled = "";

            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddCommand("get-aduser");
                    ps.AddParameter("Filter", "Name -eq " + emplid);
                    ps.AddParameter("Properties", "*");
                    Collection<PSObject> names = ps.Invoke();
                    PSObject ob = names.FirstOrDefault();
                    if (ob != null)
                    {
                        if (ob.Properties["samaccountname"].Value != null)
                            searchName = ob.Properties["samaccountname"].Value.ToString();
                        if (ob.Properties["City"].Value != null)
                            searchCity = ob.Properties["City"].Value.ToString();
                        if (ob.Properties["Country"].Value != null)
                            searchCountry = ob.Properties["Country"].Value.ToString();
                        if (ob.Properties["Department"].Value != null)
                            searchDepartment = ob.Properties["Department"].Value.ToString();
                        if (ob.Properties["Description"].Value != null)
                            searchDescription = ob.Properties["Description"].Value.ToString();
                        if (ob.Properties["DisplayName"].Value != null)
                            searchDisplayName = ob.Properties["DisplayName"].Value.ToString();
                        if (ob.Properties["EmployeeID"].Value != null)
                            searchEmployeeID = ob.Properties["EmployeeID"].Value.ToString();
                        if (ob.Properties["GivenName"].Value != null)
                            searchGivenName = ob.Properties["GivenName"].Value.ToString();
                        if (ob.Properties["OfficePhone"].Value != null)
                            searchOfficePhone = ob.Properties["OfficePhone"].Value.ToString();
                        if (ob.Properties["Initials"].Value != null)
                            searchInitials = ob.Properties["Initials"].Value.ToString();
                        if (ob.Properties["Office"].Value != null)
                            searchOffice = ob.Properties["Office"].Value.ToString();
                        if (ob.Properties["SamAccountName"].Value != null)
                            searchSamAccountName = ob.Properties["SamAccountName"].Value.ToString();
                        if (ob.Properties["State"].Value != null)
                            searchState = ob.Properties["State"].Value.ToString();
                        if (ob.Properties["StreetAddress"].Value != null)
                            searchStreetAddress = ob.Properties["StreetAddress"].Value.ToString();
                        if (ob.Properties["Surname"].Value != null)
                            searchSurname = ob.Properties["Surname"].Value.ToString();
                        if (ob.Properties["Title"].Value != null)
                            searchTitle = ob.Properties["Title"].Value.ToString();
                        if (ob.Properties["ObjectClass"].Value != null)
                            searchObjectClass = ob.Properties["ObjectClass"].Value.ToString();
                        if (ob.Properties["UserPrincipalName"].Value != null)
                            searchUserPrincipalName = ob.Properties["UserPrincipalName"].Value.ToString();
                        if (ob.Properties["Path"].Value != null)
                            searchPath = ob.Properties["Path"].Value.ToString();
                        if (ob.Properties["PostalCode"].Value != null)
                            searchPostalCode = ob.Properties["PostalCode"].Value.ToString();

                        if (ob.Properties["enabled"].Value != null)
                            searchEnabled = ob.Properties["enabled"].Value.ToString();
                        //The following lines contain a field that has not yet been implemented
                        /*if (ob.Properties["ipphone"].Value != null)
                            searchIPPhone = ob.Properties["ipphone"].Value.ToString();*/

                        ADUserModel toReturn = new ADUserModel(searchCity, searchName, searchDepartment,
                            searchDescription, searchDisplayName, searchEmployeeID, searchGivenName, searchOfficePhone,
                            searchInitials, searchOffice, searchPostalCode, searchSamAccountName, searchState,
                            searchStreetAddress, searchSurname, searchTitle, searchUserPrincipalName, searchPath, searchIPPhone,
                            searchMSExchHideFromAddressList, searchChangePasswordAtLogon, searchEnabled, searchType, "");
                        return toReturn;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// This is a driver method to be called from the MSActorController. it creates a new user in AD, and returns 
        /// the status message of the request. 
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel NewADUserDriver(ADUserModel user)
        {
            try
            {
                using (PowerShell powershell = PowerShell.Create())
                {
                    //Password nonsense to follow
                    PSCommand command = new PSCommand();
                    command.AddCommand("ConvertTo-SecureString");
                    command.AddParameter("AsPlainText");
                    command.AddParameter("String", user.accountPassword);
                    command.AddParameter("Force");
                    powershell.Commands = command;
                    Collection<PSObject> passHashCollection = powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();
                    PSObject toPass = passHashCollection.First();   //this is the password wrapped in a psobject

                    command = new PSCommand();
                    command.AddCommand("new-aduser");
                    command.AddParameter("name", user.name); //Name used to be emplid, but has since been changed
                    command.AddParameter("accountpassword", toPass);
                    command.AddParameter("changepasswordatlogon", user.changepasswordatlogon);
                    command.AddParameter("city", user.city);
                    //command.AddParameter("country", user.country);
                    command.AddParameter("department", user.department);
                    command.AddParameter("description", user.description);
                    command.AddParameter("displayname", user.displayname);
                    command.AddParameter("employeeid", user.employeeid);
                    command.AddParameter("enabled", user.enabled);
                    command.AddParameter("givenname", user.givenname);
                    command.AddParameter("officephone", user.officephone);
                    command.AddParameter("initials", user.initials);
                    command.AddParameter("office", user.office);
                    command.AddParameter("postalcode", user.postalcode);
                    command.AddParameter("samaccountname", user.samaccountname);
                    command.AddParameter("state", user.state);
                    command.AddParameter("streetaddress", user.streetaddress);
                    command.AddParameter("surname", user.surname);
                    command.AddParameter("Title", user.title);
                    command.AddParameter("type", user.type);
                    command.AddParameter("userprincipalname", user.userprincipalname);
                    command.AddParameter("path", user.path);
                    if (user.ipphone != null)
                    {
                        Hashtable attrHash = new Hashtable
                        {
                            {"ipPhone", user.ipphone }
                        };
                        command.AddParameter("OtherAttributes", attrHash);
                    }
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();
                    bool adFinished = false;
                    int count = 0;
                    String objectNotFoundMessage = "Cannot find an object with identity";
                    while (adFinished == false && count < 3)
                    {
                        command = new PSCommand();
                        command.AddCommand("get-aduser");
                        command.AddParameter("identity", user.samaccountname);
                        powershell.Commands = command;
                        Collection<PSObject> check = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            if (powershell.Streams.Error[0].Exception.Message.Contains(objectNotFoundMessage))
                            {
                                System.Threading.Thread.Sleep(1000);
                            }
                            else
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        powershell.Streams.ClearStreams();
                        if (check.FirstOrDefault() != null)
                        {
                            adFinished = true;
                        }
                        count++;
                    }

                    if(count > 2)
                    {
                        throw new Exception("Retry count exceeded. May indicate account creation issue");
                    }
                }

                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }


        /// <summary>
        /// This method changes the surname of a user in AD. 
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="samaccountname"></param>
        /// <param name="field"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel ChangeUserValueDriver(string employeeid, string samaccountname, string field, string value)
        {  
            try
            {
                if(value == "")
                {
                    value = null;
                }
                string dName;
                PSObject user = util.getADUser(employeeid, samaccountname);
                if (user == null)
                {
                    throw new Exception("User was not found.");
                }
                dName = user.Properties["DistinguishedName"].Value.ToString();
                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Set-ADUser");
                    command.AddParameter("Identity", dName);
                    if (field.ToLower() == "ipphone")
                    {
                        Hashtable attrHash = new Hashtable
                        {
                            { field, value }
                        };
                        command.AddParameter("replace", attrHash);
                    }
                    else
                    {
                        command.AddParameter(field, value);
                    }
                    command.AddParameter("ErrorVariable", "Err");
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// This method changes the home directory of a user in AD.
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="samaccountname"></param>
        /// <param name="homedirectory"></param>
        /// <param name="homedrive"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel SetHomeDirectory(string employeeid, string samaccountname, string homedirectory, string homedrive)
        {
            UtilityController util = new UtilityController();
            try
            {
                string dName;
                PSObject user = util.getADUser(employeeid, samaccountname);
                if (user == null)
                {
                    throw new Exception("User was not found.");
                }
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();

                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Set-ADUser");
                    command.AddParameter("Identity", dName);
                    command.AddParameter("homedirectory", homedirectory);
                    command.AddParameter("homedrive", homedrive);
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// This method creates a new AD group
        /// </summary>
        /// <param name="group_name"></param>
        /// <param name="group_description"></param>
        /// <param name="group_info"></param>
        /// <param name="group_ad_path"></param>
        /// <param name="group_category"></param>
        /// <param name="group_scope"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel NewADGroup(string group_name, string group_description, string group_info, string group_ad_path, string group_category, string group_scope)
        {
            UtilityController util = new UtilityController();
            try
            {
                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("New-ADGroup");
                    command.AddParameter("name", group_name);
                    command.AddParameter("description", group_description);
                    command.AddParameter("groupcategory", group_category);
                    command.AddParameter("displayname", group_info);
                    command.AddParameter("path", group_ad_path);
                    command.AddParameter("groupscope", group_scope);
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    if (group_category == "distribution")
                    {
                        ExchangeController control = new ExchangeController();
                        return control.EnableDistributionGroup(group_name);
                    }

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// This method removes an existing AD group
        /// </summary>
        /// <param name="group_identity"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel RemoveADGroup(string group_identity)
        {
            UtilityController util = new UtilityController();
            try
            {
                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Remove-ADGroup");
                    command.AddParameter("identity", group_identity);
                    command.AddParameter("confirm", false);
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Add user to an AD group
        /// </summary>
        /// <param name="group_identity"></param>
        /// <param name="group_member"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel AddADGroupMember(string group_identity, string group_member)
        {
            try
            {
                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Add-ADGroupMember");
                    command.AddParameter("identity", group_identity);
                    command.AddParameter("member", group_member);
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Remove user from an AD group
        /// </summary>
        /// <param name="group_identity"></param>
        /// <param name="group_member"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel RemoveADGroupMember(string group_identity, string group_member)
        {
            try
            {
                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Remove-ADGroupMember");
                    command.AddParameter("identity", group_identity);
                    command.AddParameter("member", group_member);
                    command.AddParameter("confirm", false);
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Set password
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="samaccountname"></param>
        /// <param name="accountpassword"></param>
        /// <param name="changepasswordatlogon"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel SetPassword(string employeeid, string samaccountname, string accountpassword, string changepasswordatlogon)
        {
            MSActorReturnMessageModel errorMessage;
            UtilityController util = new UtilityController();
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    // Try without the runspace stuff first
                    //Runspace runspace = RunspaceFactory.CreateRunspace();
                    //powershell.Runspace = runspace;
                    //runspace.Open();

                    PSObject user = util.getADUser(employeeid, samaccountname);
                    if (user == null)
                    {
                        throw new Exception("User was not found.");
                    }

                    PSCommand command = new PSCommand();
                    command.AddCommand("ConvertTo-SecureString");
                    command.AddParameter("String", accountpassword);
                    command.AddParameter("AsPlainText");
                    command.AddParameter("Force");
                    powershell.Commands = command;
                    Collection<PSObject> pwd = powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    if (pwd.Count != 1)
                    {
                        // This may not be reached anymore
                        throw new Exception("Unexpected return from creating password secure string.");
                    }

                    command = new PSCommand();
                    command.AddCommand("Set-ADAccountPassword");
                    command.AddParameter("Identity", user);
                    command.AddParameter("NewPassword", pwd[0]);
                    command.AddParameter("Reset");
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    command = new PSCommand();
                    command.AddCommand("Set-AdUser");
                    command.AddParameter("Identity", user);
                    command.AddParameter("ChangePasswordAtLogon", Boolean.Parse(changepasswordatlogon));
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// Delete entry for user
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="samaccountname"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel RemoveADObject(string employeeid, string samaccountname)
        {
            UtilityController util = new UtilityController();
            try
            {
                string dName;
                PSObject user = util.getADUser(employeeid, samaccountname);
                if (user == null)
                {
                    throw new Exception("User was not found.");
                }
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();

                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Get-ADUser");
                    command.AddParameter("Identity", dName);
                    command.AddCommand("Get-ADObject");
                    command.AddCommand("Remove-ADObject");
                    command.AddParameter("confirm", false);
                    command.AddParameter("recursive");
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// ...
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="searchbase"></param>
        /// <param name="old_samaccountname"></param>
        /// <param name="new_samaccountname"></param>
        /// <param name="userprincipalname"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel ChangeUsername(string employeeid, string old_samaccountname, string new_samaccountname, string userprincipalname)
        {
            UtilityController util = new UtilityController();
            try
            {
                // debugging:
                // $user = Get-ADUser -Filter "employeeid -eq '9999998'" -SearchBase 'OU=Accounts,DC=spudev,DC=corp' -Properties cn,displayname,givenname,initials
                // $userDN =$($user.DistinguishedName)
                // Set - ADUser - identity $userDN - sAMAccountName ‘wclinton’ -UserPrincipalName ‘wclinton @spudev.corp’  -ErrorVariable Err

                string dName;
                PSObject user = util.getADUser(employeeid, old_samaccountname);
                if (user == null)
                {
                    throw new Exception("User was not found.");
                }
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();

                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Get-ADUser");
                    command.AddParameter("Identity", dName);
                    command.AddCommand("Set-Variable");
                    command.AddParameter("Name", "user");
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    command = new PSCommand();
                    command.AddScript("$($user.DistinguishedName)");
                    command.AddCommand("Set-Variable");
                    command.AddParameter("Name", "userDN");
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    command = new PSCommand();
                    command.AddScript(String.Format("Set-ADUser -Identity $userDN -sAMAccountName {0} -UserPrincipalName {1} -ErrorVariable Err", new_samaccountname, userprincipalname));
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    command = new PSCommand();
                    command.AddScript(String.Format("Rename-ADObject -Identity $userDN -NewName {0}", new_samaccountname));
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        /// <summary>
        /// ...
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="samaccountname"></param>
        /// <param name="ipphone"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel SetIPPhone(string employeeid, string samaccountname, string ipphone)
        {
            UtilityController util = new UtilityController();
            try
            {
                string dName;
                PSObject user = util.getADUser(employeeid, samaccountname);
                if (user == null)
                {
                    throw new Exception("User was not found.");
                }
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();

                using (PowerShell powershell = PowerShell.Create())
                {
                    PSCommand command = new PSCommand();
                    command.AddCommand("Get-ADUser");
                    command.AddParameter("Identity", dName);
                    command.AddCommand("Set-ADUser");
                    if (ipphone != null)
                    {
                        Hashtable ipPhoneHash = new Hashtable
                        {
                            { "ipPhone", ipphone }
                        };
                        command.AddParameter("replace", ipPhoneHash);
                    }
                    powershell.Commands = command;
                    powershell.Invoke();
                    if (powershell.Streams.Error.Count > 0)
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                    powershell.Streams.ClearStreams();

                    MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }
        
    }
}