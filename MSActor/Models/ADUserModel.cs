using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MSActor.Models
{
    public class ADUserModel
    {
        public string emplid;
        public string identity;
        public string city;
        public string country;
        public string department;
        public string description;
        public string displayname;
        public string employeeid;
        
        public string givenname;
        public string officephone;
        public string initials;
        public string office;
        public string postalcode;
        public string samaccountname;
        public string state;
        public string streetaddress;
        public string surname;
        public string title;
        public string type;
        public string userprincipalname;
        public string path;


        public ADUserModel(string city, string country, string department, string description, 
            string displayname, string employeeid, string givenname, string officephone, string initials, 
            string office, string postalcode, string samaccountname, string state, string streetaddress, 
            string surname, string title, string userprincipalname, string path)
        {
            this.emplid = employeeid;
            this.city = city;
            this.country = country;
            this.department = department;
            this.description = description;
            this.displayname = displayname;
            this.employeeid = employeeid;
            this.givenname = givenname;
            this.officephone = officephone;
            this.initials = initials;
            this.office = office;
            this.postalcode = postalcode;
            this.samaccountname = samaccountname;
            this.state = state;
            this.streetaddress = streetaddress;
            this.surname = surname;
            this.title = title;
            this.userprincipalname = userprincipalname;
            this.path = path;
        }


        public string ToString()
        {
            string toReturn = "";
            toReturn += "emplid = " + emplid + ", ";
            toReturn += "city = " + city + ", ";
            toReturn += "country = " + country + ", ";
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