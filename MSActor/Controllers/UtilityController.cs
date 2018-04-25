using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Web;

namespace MSActor.Controllers
{
    public class UtilityController
    {
        public PSObject getADUser(string employeeid, string samaccountname)
        {
            PowerShell ps = PowerShell.Create();
            string query = "get-aduser -filter {employeeid -eq " + employeeid + " -and samaccountname -eq \"" + samaccountname + "\"}";
            
            ps.AddScript(query);
            Collection<PSObject> users = ps.Invoke();
            
            foreach(PSObject ob in users)
            {
                return ob;
            }
            return null;
        }

        public Runspace ConnectRemotePSSession(string serverpath)
        {
            PSSessionOption option = new PSSessionOption();
            string url = serverpath;
            System.Uri uri = new Uri(url);

            Runspace runspace = RunspaceFactory.CreateRunspace();

            PowerShell powershell = PowerShell.Create();
            PSCommand command = new PSCommand();
            command.AddCommand("New-PSSession");
            command.AddParameter("ConfigurationName", "Microsoft.Exchange");
            command.AddParameter("ConnectionUri", uri);
            command.AddParameter("Authentication", "Default");
            powershell.Commands = command;
            runspace.Open();
            powershell.Runspace = runspace;
            Collection<PSSession> result = powershell.Invoke<PSSession>();

            powershell = PowerShell.Create();
            command = new PSCommand();
            command.AddCommand("Set-Variable");
            command.AddParameter("Name", "ra");
            command.AddParameter("Value", result[0]);

            powershell.Commands = command;
            powershell.Runspace = runspace;
            powershell.Invoke();

            powershell = PowerShell.Create();
            command = new PSCommand();
            command.AddScript("Import-PSSession -Session $ra");
            powershell.Commands = command;
            powershell.Runspace = runspace;
            powershell.Invoke();
            Debug.WriteLine("About to make it out of the util controller");
            return runspace;
        }
    }
}