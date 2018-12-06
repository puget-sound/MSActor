using MSActor.Controllers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ADUserModel
    {
        
        public string name = null;
        public string city = null;
        //public string country = null;
        public string department = null;
        public string description = null;
        public string displayname = null;
        public string employeeid = null;
        public string givenname = null;
        public string officephone = null;
        public string initials = null;
        public string office = null;
        public string postalcode = null;
        public string samaccountname = null;
        public string state = null;
        public string streetaddress = null;
        public string surname = null;
        public string title = null;
        public string type = null;
        public string userprincipalname = null;
        public string path = null;
        public string ipphone = null;
        public string msExchHideFromAddressList = null;
        public bool changepasswordatlogon = false;
        
        public bool enabled = false;
        public string accountPassword;

        
        public ADUserModel(string city, string name, string department, string description, 
            string displayname, string employeeid, string givenname, string officephone, string initials, 
            string office, string postalcode, string samaccountname, string state, string streetaddress, 
            string surname, string title, string userprincipalname, string path, string ipphone, 
            string msExchHideFromAddressList, string changepasswordatlogon, string enabled, string type, string accountPassword)
        {
            string LogMessage = "NewADUser|name: " + name + "|city: " + city + "|department: " + department + "|description: " + description + "|" +
            "displayname: " + displayname + "|employeeid: " + employeeid + "|givenname: " + givenname + "|officephone: " + officephone +
            "|initials: " + initials + "|office: " + office + "|postalcode: " + postalcode + "|samaccountname: " + samaccountname + "|state: " + state
            + "|streetaddress: " + streetaddress + "|surname: " + surname + "|title: " + title + "|type: " + type + "|userprincipalname: " +
            userprincipalname + "|path: " + path + "|ipphone: " + ipphone + "|msExchHideFromAddressList: " + msExchHideFromAddressList +
            "|changepasswordatlogon: " + changepasswordatlogon + "|enabled: " + enabled;
            UtilityController util = new UtilityController();
            util.LogMessage(LogMessage);
            if (name != "")
                this.name = name;      
            if (city != "")
                this.city = city;
            /*if (country != "")
                this.country = country;*/
            if (department != "")
                this.department = department;
            if (description != "")
                this.description = description;
            if (displayname != "")
                this.displayname = displayname;
            if (employeeid != "")
                this.employeeid = employeeid;           
            if (givenname != "")
                this.givenname = givenname;            
            if (officephone != "")           
                this.officephone = officephone;            
            if (initials != "")            
                this.initials = initials;           
            if (office != "")            
                this.office = office;            
            if (postalcode != "")            
                this.postalcode = postalcode;            
            if (samaccountname != "")            
                this.samaccountname = samaccountname;            
            if (state != "")            
                this.state = state;            
            if (streetaddress != "")            
                this.streetaddress = streetaddress;           
            if (surname != "")            
                this.surname = surname;           
            if (title != "")            
                this.title = title;            
            if (userprincipalname != "")            
                this.userprincipalname = userprincipalname;            
            if (path != "")           
                this.path = path;
            if (ipphone != "")
                this.ipphone = ipphone;
            if (msExchHideFromAddressList != "")
                this.msExchHideFromAddressList = msExchHideFromAddressList;
            
            if (changepasswordatlogon != "")
            {
                
                if (changepasswordatlogon.ToLower() == "true")
                {
                    this.changepasswordatlogon = true;
                }
                else if (changepasswordatlogon.ToLower() == "false")
                {
                    this.changepasswordatlogon = false;
                }
            }
            
            if (enabled != "")
            {
                if (enabled.ToLower() == "true")
                {
                    this.enabled = true;
                }
                else if(enabled.ToLower() == "false")
                {
                    this.enabled = false;
                }
            }
            if (accountPassword != "")
                this.accountPassword = accountPassword;
            if (type != "")
                this.type = type;

            
        }


        public override string ToString()
        {
            string toReturn = "";
            
            toReturn += "city = " + city + ", ";
            toReturn += "department = " + department + ", ";
            toReturn += "description = " + description + ", ";
            toReturn += "displayname = " + displayname + ", ";
            toReturn += "employeeid = " + employeeid + ", ";
            toReturn += "givenname = " + givenname + ", ";
            toReturn += "officephone = " + officephone + ", ";
            toReturn += "initials = " + initials + ", ";
            toReturn += "office = " + office + ", ";
            toReturn += "postalcode = " + postalcode + ", ";
            toReturn += "samaccountname = " + samaccountname + ", ";
            toReturn += "state = " + state + ", ";
            toReturn += "streetaddress = " + streetaddress + ", ";
            toReturn += "surname = " + surname + ", ";
            toReturn += "title = " + title + ", ";
            toReturn += "userprincipalname = " + userprincipalname + ", ";
            return toReturn;
        }
    }
}