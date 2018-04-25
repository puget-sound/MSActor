using MSActor.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Web;

namespace MSActor.Controllers
{
    /// <summary>
    /// A class to hold all the File Server related methods
    /// </summary>
    public class FileServerController
    {
        public const string SuccessCode = "CMP";
        public const string ErrorCode = "ERR";

        UtilityController util;
        public FileServerController()
        {
            util = new UtilityController();
        }

        public MSActorReturnMessageModel NewDirectory(string path)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = "http://spufs01/powershell/ ";
                System.Uri uri = new Uri(url);

                Runspace runspace = RunspaceFactory.CreateRunspace();

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("New-PSSession");
                command.AddParameter("ComputerName", "spufs01.spudev.corp");
                command.AddParameter("Authentication", "Default");
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSSession> result = powershell.Invoke<PSSession>();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "s");
                command.AddParameter("Value", result[0]);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("enter-PSSession $s");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();
                
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("new-item");
                command.AddParameter("itemtype", "directory");
                Debug.WriteLine("Path is : " + path);
                command.AddParameter("path", path);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                Collection<PSObject> returns = powershell.Invoke();
                Debug.WriteLine("Boy for your sake I hope this is a 1: " + returns.Count);

                /*Debug.WriteLine("Ok it's the removal that's hosing me");
                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("remove-pssession $s");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();*/

                Debug.WriteLine("end of create folder");
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch(Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, "");
                Debug.WriteLine("First of all, you're dumb and ugly: " + e.Message);
                return errorMessage;
            }
            
        }

        public MSActorReturnMessageModel UserFolderAddAccessDriver(string path, string samaccountname)
        {
            try
            {

                Runspace runspace = util.ConnectRemotePSSession("http://spufs01/powershell/ ");

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("get-acl");
                command.AddParameter("path", path);
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSSession> result = powershell.Invoke<PSSession>();
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "acl");
                command.AddParameter("Value", result[0]);  
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("new-object system.security.accesscontrol.filesystemaccessrule(" +
                    "\"" + samaccountname + "\",\"FullControl\",\"ContainerInherit, ObjectInherit\",\"None\",\"Allow\")");
                
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSSession> permCol = powershell.Invoke<PSSession>();
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "perms");
                command.AddParameter("Value", permCol[0]);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("$acl.setaccessrule($perms)");
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke<PSSession>();
                powershell.Invoke();

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("set-acl -path " + path + " -aclobject $acl");
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke<PSSession>();
                powershell.Invoke();

                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, "");
                return errorMessage;
            }
        }
    }
}