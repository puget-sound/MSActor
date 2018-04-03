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
using ContactManager.Models;
using ContactManager.Services;
using Newtonsoft.Json;


namespace ContactManager.Controllers
{
    public class ContactController : ApiController
    {
        private ContactRepository contactRepository;
        

        public ContactController()
        {
            this.contactRepository = new ContactRepository();
        }
        /// <summary>$
        /// This is a method for creating new users in AD
        /// </summary>
        [Route("newaduser")]
        [HttpPost]
        private void NewADUser([FromBody] ADUser newUser)
        {
            var ctx = HttpContext.Current;
            var request = ctx.Request;
            Debug.WriteLine(newUser.toString());
            PowerShell ps = PowerShell.Create();    //Password nonsense to follow
            ps.AddCommand("ConvertTo-SecureString");
            ps.AddParameter("AsPlainText");
            ps.AddParameter("String", "First123!!!");
            ps.AddParameter("Force");
            Collection<PSObject> passHashCollection = ps.Invoke();
            PSObject toPass = passHashCollection.First();   //this is the password wrapped in a psobject

            ps = PowerShell.Create();
            ps.AddCommand("new-aduser");
            Debug.WriteLine("newUser.emplid = " + newUser.emplid);
            ps.AddParameter("name", newUser.emplid); //Name is actually emplid pending future change
            ps.AddParameter("accountpassword", toPass);
            ps.AddParameter("changepasswordatlogon", true);
            ps.AddParameter("city", newUser.city);
            ps.AddParameter("country", newUser.country);
            ps.AddParameter("department", newUser.department);
            ps.AddParameter("description", newUser.description);
            ps.AddParameter("displayname", newUser.displayname);
            ps.AddParameter("employeeid", newUser.emplid);
            ps.AddParameter("enabled", true);
            ps.AddParameter("givenname", newUser.givenname);
            ps.AddParameter("officephone", newUser.officephone);
            ps.AddParameter("initials", newUser.initials);
            ps.AddParameter("office", newUser.office);
            ps.AddParameter("postalcode", newUser.postalcode);
            ps.AddParameter("samaccountname", newUser.samaccountname);
            ps.AddParameter("state", newUser.state);
            ps.AddParameter("streetaddress", newUser.streetaddress);
            ps.AddParameter("surname", newUser.surname);
            ps.AddParameter("Title", newUser.title);
            ps.AddParameter("type", "user");
            ps.AddParameter("userprincipalname", newUser.userprincipalname);
            ps.AddParameter("path", newUser.path);
            Collection<PSObject> names = ps.Invoke();
           
        }
        ///<summary>
        ///this is a test for get requests
        /// </summary>
        [Route("getrequesttest")]
        [HttpGet]
        public string GetRequestTest()
        {
            return "this get request test was successful";
        }
        /// <summary>
        /// This is a method to return the information about a user based on emplid. 
        /// Parameter is a Json object with the form {emplid = XXXXXXXXX}
        /// </summary>
        [Route("getaduserbyemplid")]
        [HttpPost]
        public ADUser GetAdUserByEmplid([FromBody] EmplidWrapper emplidWrap)
        {
                try
                {
                    /*string date = DateTime.Now.Month.ToString() + "-"
                        + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Hour.ToString() + "."
                        + DateTime.Now.Minute.ToString();
                    string logFileName = date + "GetADUserByEmplid.txt";
                    string filepath = System.IO.Path.Combine("C:\\inetpub\\logs\\LogFiles\\", logFileName);
                    Debug.WriteLine("filepath is " + filepath);
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(filepath))
                    {
                        file.WriteLine("Now beginning GetADUserByEmplid");
                    }*/
                    /*var ctx = HttpContext.Current;
                    var request = ctx.Request;
                    var json = request.Form[null];
                    foreach(string s in request.Form.Keys)
                    {
                    Debug.WriteLine("Here's a key:" + s + "there is no key");
                    }
                    ADUser user = JsonConvert.DeserializeObject<ADUser>(json);
                    */
                    PowerShell ps = PowerShell.Create();
                    ps.AddCommand("get-aduser");
                    ps.AddParameter("Filter", "Name -eq " + emplidWrap.emplid);
                    ps.AddParameter("Properties", "*");
                    Collection<PSObject> names = ps.Invoke();
                    foreach (PSObject ob in names)
                    {
                        /*using (System.IO.StreamWriter file = System.IO.File.AppendText(filepath))
                        {
                            file.WriteLine("Found a thing in the foreach loop");
                        }*/
                        ADUser toReturn = new ADUser(emplidWrap.emplid);
                        if (ob.Properties["samaccountname"].Value != null)
                            toReturn.identity = ob.Properties["samaccountname"].Value.ToString();
                        if (ob.Properties["City"].Value != null)
                            toReturn.city = ob.Properties["City"].Value.ToString();
                        if (ob.Properties["Country"].Value != null)
                            toReturn.country = ob.Properties["Country"].Value.ToString();
                        if (ob.Properties["Department"].Value != null)
                            toReturn.department = ob.Properties["Department"].Value.ToString();
                        if (ob.Properties["Description"].Value != null)
                            toReturn.description = ob.Properties["Description"].Value.ToString();
                        if (ob.Properties["DisplayName"].Value != null)
                            toReturn.displayname = ob.Properties["DisplayName"].Value.ToString();
                        if (ob.Properties["EmployeeID"].Value != null)
                            toReturn.employeeid = ob.Properties["EmployeeID"].Value.ToString();
                        if (ob.Properties["GivenName"].Value != null)
                            toReturn.givenname = ob.Properties["GivenName"].Value.ToString();
                        if (ob.Properties["OfficePhone"].Value != null)
                            toReturn.officephone = ob.Properties["OfficePhone"].Value.ToString();
                        if (ob.Properties["Initials"].Value != null)
                            toReturn.initials = ob.Properties["Initials"].Value.ToString();
                        if (ob.Properties["Office"].Value != null)
                            toReturn.office = ob.Properties["Office"].Value.ToString();
                        if (ob.Properties["SamAccountName"].Value != null)
                            toReturn.samaccountname = ob.Properties["SamAccountName"].Value.ToString();
                        if (ob.Properties["State"].Value != null)
                            toReturn.state = ob.Properties["State"].Value.ToString();
                        if (ob.Properties["StreetAddress"].Value != null)
                            toReturn.streetaddress = ob.Properties["StreetAddress"].Value.ToString();
                        if (ob.Properties["Surname"].Value != null)
                            toReturn.surname = ob.Properties["Surname"].Value.ToString();
                        if (ob.Properties["Title"].Value != null)
                            toReturn.title = ob.Properties["Title"].Value.ToString();
                        if (ob.Properties["ObjectClass"].Value != null)
                            toReturn.type = ob.Properties["ObjectClass"].Value.ToString();
                        if (ob.Properties["UserPrincipalName"].Value != null)
                            toReturn.userprincipalname = ob.Properties["UserPrincipalName"].Value.ToString();
                        if (ob.Properties["Path"].Value != null)
                            toReturn.path = ob.Properties["Path"].Value.ToString();

                        return toReturn;
                    }
                    /*using (System.IO.StreamWriter file = System.IO.File.AppendText(filepath))
                    {
                        file.Write("Danger Will Robinson");
                    }*/
                    
                    
                    return null;
                }
                catch (Exception e)
                {
                    /*string filedate = DateTime.Now.Month.ToString() + "-"
                        + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Hour.ToString()
                        + "." + DateTime.Now.Minute.ToString();
                    string errorLogName = filedate + "GetADUserByEmplidErrorLog.txt";
                    string logpath = System.IO.Path.Combine("C:\\inetpub\\logs\\LogFiles\\", errorLogName);
                    
                    using (System.IO.StreamWriter file =
                        new System.IO.StreamWriter(logpath))
                    {
                        file.WriteLine(e.Message.ToString());
                    }*/
                    JsonErrorWrapper toReturn = new JsonErrorWrapper(e.Message);
                    return null;

                }
        }
    }
}
