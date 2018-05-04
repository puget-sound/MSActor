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
using System.Text.RegularExpressions;
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

        public MSActorReturnMessageModel NewDirectory(string computername, string path)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = String.Format("http://{0}:5985/wsman", computername);
                Uri uri = new Uri(url);
                WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                Runspace runspace = RunspaceFactory.CreateRunspace(conn);
                runspace.Open();

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("New-Item");
                command.AddParameter("ItemType", "directory");
                command.AddParameter("Path", path);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                Collection<PSObject> returns = powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

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

        public MSActorReturnMessageModel RemoveDirectory(string computername, string path)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = String.Format("http://{0}:5985/wsman", computername);
                Uri uri = new Uri(url);
                WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                Runspace runspace = RunspaceFactory.CreateRunspace(conn);
                runspace.Open();

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("Remove-Item");
                command.AddParameter("Path", path);
                command.AddParameter("Recurse");
                command.AddParameter("Force");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();

                if (powershell.Streams.Error.Count > 0)
                {
                    RemoteException ex = powershell.Streams.Error[0].Exception as RemoteException;
                    if (ex.SerializedRemoteException.TypeNames.Contains("Deserialized.System.IO.PathTooLongException"))
                    {
                        PowerShell powershell1 = PowerShell.Create();
                        PSCommand command1 = new PSCommand();
                        command1.AddCommand("Set-ExecutionPolicy");
                        command1.AddParameter("ExecutionPolicy", "RemoteSigned");
                        command1.AddParameter("Scope", "Process");
                        command1.AddParameter("Force");
                        powershell1.Commands = command1;
                        powershell1.Runspace = runspace;
                        powershell1.Invoke();
                        if (powershell1.Streams.Error.Count > 0)
                        {
                            throw powershell1.Streams.Error[0].Exception;
                        }

                        command1 = new PSCommand();
                        command1.AddScript(". D:\\PathTooLong.ps1");
                        powershell1.Commands = command1;
                        powershell1.Runspace = runspace;
                        powershell1.Invoke();
                        if (powershell1.Streams.Error.Count > 0)
                        {
                            throw powershell1.Streams.Error[0].Exception;
                        }

                        command1 = new PSCommand();
                        command1.AddCommand("Remove-PathToLongDirectory");
                        command1.AddArgument(path);
                        powershell1.Commands = command1;
                        powershell1.Runspace = runspace;
                        powershell1.Invoke();
                        if (powershell1.Streams.Error.Count > 0)
                        {
                            throw powershell1.Streams.Error[0].Exception;
                        }
                    }
                    else
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                }

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

        public MSActorReturnMessageModel AddNetShare(string name, string computername, string path)
        {
            MSActorReturnMessageModel successMessage;

            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = String.Format("http://{0}:5985/wsman", computername);
                Uri uri = new Uri(url);
                WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                Runspace runspace = RunspaceFactory.CreateRunspace(conn);
                runspace.Open();

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                string script = String.Format("net share {0}={1} \"/GRANT:Everyone,Full\"", name, path);
                command.AddScript(script);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    if (powershell.Streams.Error[0].FullyQualifiedErrorId == "NativeCommandError")
                    {
                        // If the share already exists it might be okay (see "else" below). Otherwise this is an error.
                        if (powershell.Streams.Error[0].Exception.Message != "The name has already been shared.")
                        {
                            System.Text.StringBuilder msgBuilder = new System.Text.StringBuilder();
                            foreach (ErrorRecord errorRec in powershell.Streams.Error)
                            {
                                // Kludge to fix a weird bug with blank lines in the error output
                                if (errorRec.CategoryInfo.ToString() == errorRec.Exception.Message)
                                {
                                    msgBuilder.AppendLine();
                                }
                                else
                                {
                                    msgBuilder.AppendLine(errorRec.Exception.Message);
                                }
                            }
                            throw new Exception(msgBuilder.ToString());
                        }
                        else
                        {
                            // Check that the existing share has the same path
                            command = new PSCommand();
                            script = String.Format("net share {0}", name);
                            command.AddScript(script);
                            powershell.Commands = command;
                            powershell.Runspace = runspace;
                            Collection<PSObject> ret = powershell.Invoke();
                            // Find the first (hopefully only) line in the output with "Path" at the beginning
                            string pathResultLine = ret.First(x => (x.BaseObject as string).StartsWith("Path")).BaseObject as string;
                            // Separate the line into (Non-words)[bunch of spaces](Rest of line)
                            GroupCollection groups = Regex.Match(pathResultLine, @"(\S+)\s+(.+)$").Groups;
                            // The (Rest of line) part of the separation is the path value
                            string pathResult = groups[2].Value;
                            if (pathResult == path)
                            {
                                successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                                return successMessage;
                            }
                            else
                            {
                                throw new Exception(String.Format("Share '{0}' already exists for a different path '{1}'.", name, pathResult));
                            }
                        }
                    }
                    else
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                }

                successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;

            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("ERROR: " + e.Message);
                return errorMessage;
            }
        }

        public MSActorReturnMessageModel RemoveNetShare(string name, string computername, string path)
        {
            MSActorReturnMessageModel successMessage;

            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = String.Format("http://{0}:5985/wsman", computername);
                Uri uri = new Uri(url);
                WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                Runspace runspace = RunspaceFactory.CreateRunspace(conn);
                runspace.Open();

                // First check that the share name is for the correct path
                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                string script = String.Format("net share {0}", name);
                command.AddScript(script);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                Collection<PSObject> ret = powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    if (powershell.Streams.Error[0].FullyQualifiedErrorId == "NativeCommandError")
                    {
                        // If the share does not exist return success
                        if (powershell.Streams.Error[0].Exception.Message == "This shared resource does not exist.")
                        {
                            successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                            return successMessage;
                        }
                        else
                        {
                            System.Text.StringBuilder msgBuilder = new System.Text.StringBuilder();
                            foreach (ErrorRecord errorRec in powershell.Streams.Error)
                            {
                                // Kludge to fix a weird bug with blank lines in the error output
                                if (errorRec.CategoryInfo.ToString() == errorRec.Exception.Message)
                                {
                                    msgBuilder.AppendLine();
                                }
                                else
                                {
                                    msgBuilder.AppendLine(errorRec.Exception.Message);
                                }
                            }
                            throw new Exception(msgBuilder.ToString());
                        }
                    }
                    else
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                }

                // Find the first (hopefully only) line in the output with "Path" at the beginning
                string pathResultLine = ret.First(x => (x.BaseObject as string).StartsWith("Path")).BaseObject as string;
                // Separate the line into (Non-words)[bunch of spaces](Rest of line)
                GroupCollection groups = Regex.Match(pathResultLine, @"(\S+)\s+(.+)$").Groups;
                // The (Rest of line) part of the separation is the path value
                string existingPath = groups[2].Value;
                if (path != existingPath)
                {
                    throw new Exception(String.Format("Share '{0}' is for path '{1}', different than specified.", name, existingPath));
                }

                // Now delete the share
                script = String.Format("net share {0} /delete", name);
                command.AddScript(script);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    if (powershell.Streams.Error[0].FullyQualifiedErrorId == "NativeCommandError")
                    {
                        System.Text.StringBuilder msgBuilder = new System.Text.StringBuilder();
                        foreach (ErrorRecord errorRec in powershell.Streams.Error)
                        {
                            // Kludge to fix a weird bug with blank lines in the error output
                            if (errorRec.CategoryInfo.ToString() == errorRec.Exception.Message)
                            {
                                msgBuilder.AppendLine();
                            }
                            else
                            {
                                msgBuilder.AppendLine(errorRec.Exception.Message);
                            }
                        }
                        throw new Exception(msgBuilder.ToString());
                    }
                    else
                    {
                        throw powershell.Streams.Error[0].Exception;
                    }
                }

                successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                Debug.WriteLine("ERROR: " + e.Message);
                return errorMessage;
            }
        }

        public MSActorReturnMessageModel AddUserFolderAccess(string computername, string path, string samaccountname)
        {
            try
            {

                //Runspace runspace = util.ConnectRemotePSSession("http://spufs01/powershell/ ");

                PSSessionOption option = new PSSessionOption();
                string url = "http://spufs01:5985/wsman";
                Uri uri = new Uri(url);
                WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                Runspace runspace = RunspaceFactory.CreateRunspace(conn);
                runspace.Open();

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("get-acl");
                command.AddParameter("path", path);
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSSession> result = powershell.Invoke<PSSession>();
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "acl");
                command.AddParameter("Value", result[0]);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("new-object system.security.accesscontrol.filesystemaccessrule(" +
                    "\"" + samaccountname + "\",\"FullControl\",\"ContainerInherit, ObjectInherit\",\"None\",\"Allow\")");
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                Collection<PSSession> permCol = powershell.Invoke<PSSession>();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "perms");
                command.AddParameter("Value", permCol[0]);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("$acl.setaccessrule($perms)");
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                powershell = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("set-acl -path " + path + " -aclobject $acl");
                powershell.Commands = command;
                runspace.Open();
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                return successMessage;
            }
            catch (Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                return errorMessage;
            }
        }
    }
}