using MSActor.Models;
using System;
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
    public class ExchangeController
    {
        public const string SuccessCode = "CMP";
        public const string ErrorCode = "ERR";
        public ExchangeController()
        {

        }
        public MSActorReturnMessageModel EnableMailboxDriver(string database, string alias, string emailaddresses)
        {   
            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = "http://spudevexch13a.spudev.corp/powershell/";
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
                
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Enable-Mailbox");
                command.AddParameter("identity", alias);
                
                command.AddParameter("database", database);
                command.AddParameter("alias", alias);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("set-mailbox");
                command.AddParameter("identity", alias);
                command.AddParameter("emailaddresses", emailaddresses);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
                
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("ERROR: " + e.Message);
                return errorMessage;
            }
        }

        public MSActorReturnMessageModel RemoveMailboxDriver(string employeeid, string samaccountname)
        {
            PowerShell ps = PowerShell.Create();
            ps.AddScript("get-aduser –filter * -properties * | where-object {$_.employeeid –eq }");
            return null;
        }
    }

}