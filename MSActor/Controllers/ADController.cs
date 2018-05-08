/*
using MSActor.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Web;
*/
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
                PowerShell ps = PowerShell.Create();
                ps.AddCommand("get-aduser");
                ps.AddParameter("Filter", "Name -eq " + emplid);
                ps.AddParameter("Properties", "*");
                Collection<PSObject> names = ps.Invoke();
                foreach (PSObject ob in names)
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
                return null;
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
                PowerShell ps = PowerShell.Create();    //Password nonsense to follow
                ps.AddCommand("ConvertTo-SecureString");
                ps.AddParameter("AsPlainText");
                ps.AddParameter("String", user.accountPassword);
                ps.AddParameter("Force");
                Collection<PSObject> passHashCollection = ps.Invoke();
                PSObject toPass = passHashCollection.First();   //this is the password wrapped in a psobject

                ps = PowerShell.Create();
                ps.AddCommand("new-aduser");
                ps.AddParameter("name", user.name); //Name used to be emplid, but has since been changed
                ps.AddParameter("accountpassword", toPass);
                ps.AddParameter("changepasswordatlogon", user.changepasswordatlogon);
                ps.AddParameter("city", user.city);
                //ps.AddParameter("country", user.country);
                ps.AddParameter("department", user.department);
                ps.AddParameter("description", user.description);
                ps.AddParameter("displayname", user.displayname);
                ps.AddParameter("employeeid", user.employeeid);
                ps.AddParameter("enabled", user.enabled);
                ps.AddParameter("givenname", user.givenname);
                ps.AddParameter("officephone", user.officephone);
                ps.AddParameter("initials", user.initials);
                ps.AddParameter("office", user.office);
                ps.AddParameter("postalcode", user.postalcode);
                ps.AddParameter("samaccountname", user.samaccountname);
                ps.AddParameter("state", user.state);
                ps.AddParameter("streetaddress", user.streetaddress);
                ps.AddParameter("surname", user.surname);
                ps.AddParameter("Title", user.title);
                ps.AddParameter("type", user.type);
                ps.AddParameter("userprincipalname", user.userprincipalname);
                ps.AddParameter("path", user.path);
                //ps.AddParameter("ipphone", user.ipphone);
                Collection<PSObject> names = ps.Invoke();

            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Error: " + e.Message);
                return errorMessage;
            }
            MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
            return successMessage;
        }


        /// <summary>
        /// This method changes the surname of a user in AD. 
        /// </summary>
        /// <param name="employeeid"></param>
        /// <param name="field"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public MSActorReturnMessageModel ChangeUserValueDriver(string employeeid, string samaccountname, string field, string value)
        {
            UtilityController util = new UtilityController();
            try
            {
                string dName;
                PSObject user = util.getADUser(employeeid, samaccountname);
                dName = user.Properties["DistinguishedName"].Value.ToString();
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Set-ADUser");
                ex.AddParameter("Identity", dName);
                ex.AddParameter(field, value);
                ex.AddParameter("ErrorVariable", "Err");
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Set-ADUser");
                ex.AddParameter("Identity", dName);
                ex.AddParameter("homedirectory", homedirectory);
                ex.AddParameter("homedrive", homedrive);
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("New-ADGroup");
                ex.AddParameter("name", group_name);
                ex.AddParameter("description", group_description);
                ex.AddParameter("groupcategory", group_category);
                ex.AddParameter("displayname", group_info);
                ex.AddParameter("path", group_ad_path);
                ex.AddParameter("groupscope", group_scope);
                if (group_category == "distribution")
                {
                    ex.AddStatement();
                    ex.AddCommand("enable-distributiongroup");
                    ex.AddParameter("identity", group_name);
                }
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Remove-ADGroup");
                ex.AddParameter("identity", group_identity);
                ex.AddParameter("confirm", false);
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
            UtilityController util = new UtilityController();
            try
            {
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Add-ADGroupMember");
                ex.AddParameter("identity", group_identity);
                ex.AddParameter("member", group_member);
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
            UtilityController util = new UtilityController();
            try
            {
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Remove-ADGroupMember");
                ex.AddParameter("identity", group_identity);
                ex.AddParameter("member", group_member);
                ex.AddParameter("confirm", false);
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
            UtilityController util = new UtilityController();
            try
            {
                PowerShell ex = PowerShell.Create();
                //ex.AddCommand("");
                //ex.AddParameter("", employeeid);
                //ex.AddParameter("", samaccountname);
                //ex.AddParameter("", accountpassword);
                //ex.AddParameter("", changepasswordatlogon);
                //ex.Invoke();
                // debugging
                //dsmod user –empid 9999998 -pwd Saxman123! -mustchpwd no
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Get-ADUser");
                ex.AddParameter("Identity", dName);
                ex.AddCommand("Get-ADObject");
                ex.AddCommand("Remove-ADObject");
                ex.AddParameter("confirm", false);
                ex.AddParameter("recursive");
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
            }
        }

        public MSActorReturnMessageModel ChangeUsername(string employeeid, string searchbase, string samaccountname, string userprincipalname)
        {
            UtilityController util = new UtilityController();
            try
            {
                //$user = Get-ADUser -Filter "employeeid -eq '9999998'" -SearchBase 'OU=Accounts,DC=spudev,DC=corp' -Properties cn,displayname,givenname,initials
                //$userDN =$($user.DistinguishedName)
                //Set - ADUser - identity $userDN - sAMAccountName ‘wclinton’ -UserPrincipalName ‘wclinton @spudev.corp’  -ErrorVariable Err

                string dName;
                PSObject user = util.getADUser(employeeid, samaccountname);
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Get-ADUser");
                ex.AddParameter("Identity", dName);
                ex.AddParameter("SearchBase", searchbase);
                // debugging
                //ex.AddParameter("Properties", );
                ex.AddCommand("Set-ADUser");
                ex.AddParameter("Identity", dName);
                ex.AddParameter("sAMAccountName", samaccountname);
                ex.AddParameter("UserPrincipalName", userprincipalname);
                ex.AddParameter("ErrorVariable", "Err");
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
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
                Debug.WriteLine(user);
                dName = user.Properties["DistinguishedName"].Value.ToString();
                PowerShell ex = PowerShell.Create();
                ex.AddCommand("Get-ADUser");
                ex.AddParameter("Identity", dName);
                ex.AddCommand("Set-ADUser");
                if (ipphone != null)
                {
                    Hashtable ipPhoneHash = new Hashtable
                    {
                        { "ipPhone", ipphone }
                    };
                    ex.AddParameter("replace", ipPhoneHash);
                }
                ex.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                return errorMessage;
            }
        }

        /*
                public MSActorReturnMessageModel Set_msExchHideFromAddressLists(string employeeid, string samaccountname, string privacyrestriction)
                {
                    UtilityController util = new UtilityController();
                    try
                    {
                        // translate "true" / "false" string to boolean value

                        string dName;
                        PSObject user = util.getADUser(employeeid, samaccountname);
                        Debug.WriteLine(user);
                        dName = user.Properties["DistinguishedName"].Value.ToString();
                        PowerShell ex = PowerShell.Create();
                        ex.AddCommand("Get-ADUser");
                        ex.AddParameter("Identity", dName);
                        // debugging
                        //Get-mailbox $user.samaccountname | set-mailbox –HiddenFromAddressListsEnabled $True
                        ex.AddCommand("Get-mailbox");
                        ex.AddParameter("Identity", dName);
                        ex.AddParameter("sAMAccountName", samaccountname);
                        ex.AddCommand("Set-mailbox");
                        ex.AddParameter("HiddenFromAddressListsEnabled", (privacyrestriction == "true") ? true : false);
                        ex.Invoke();
                        MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                        return successMessage;
                    }
                    catch (Exception e)
                    {
                        MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                        Debug.WriteLine("Ruh Roh Raggy: " + e.Message);
                        return errorMessage;
                    }
                }
        */
    }
}