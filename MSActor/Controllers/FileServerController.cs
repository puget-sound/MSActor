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
using System.Security;
using System.Security.AccessControl;
using System.Text;
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
                            StringBuilder msgBuilder = new StringBuilder();
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
                            StringBuilder msgBuilder = new StringBuilder();
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
                        StringBuilder msgBuilder = new StringBuilder();
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

        public MSActorReturnMessageModel AddUserFolderAccess(string employeeid, string samaccountname, string computername, string path, string accesstype)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                string url = String.Format("http://{0}:5985/wsman", computername);
                Uri uri = new Uri(url);
                WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                Runspace runspace = RunspaceFactory.CreateRunspace(conn);
                runspace.Open();

                PSObject user = util.getADUser(employeeid, samaccountname);
                string identity = user.Properties["SamAccountName"].Value as string;

                PowerShell powershell = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("Get-Acl");
                command.AddParameter("Path", path);
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "acl");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                Collection<PSObject> result = powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                command = new PSCommand();
                command.AddCommand("New-Object");
                command.AddParameter("TypeName", "System.Security.AccessControl.FileSystemAccessRule");
                command.AddParameter("ArgumentList",
                    new object[]
                    {
                        identity,
                        accesstype,
                        "ContainerInherit,ObjectInherit",
                        "None",
                        "Allow"
                    });
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "perms");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                result = powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                command = new PSCommand();
                command.AddScript("$acl.SetAccessRule($perms)");
                powershell.Commands = command;
                powershell.Runspace = runspace;
                powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    throw powershell.Streams.Error[0].Exception;
                }

                command = new PSCommand();
                command.AddScript(String.Format("Set-Acl -AclObject $acl -Path {0}", path));
                powershell.Commands = command;
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

        public MSActorReturnMessageModel AddDirQuota(string computername, string path, string limit)
        {
            MSActorReturnMessageModel errorMessage, successMessage;
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
                string script = String.Format("dirquota quota add /path:\"{0}\" /limit:{1}", path, limit);
                command.AddScript(script);
                powershell.Commands = command;
                powershell.Runspace = runspace;
                Collection<PSObject> ret = powershell.Invoke();
                if (powershell.Streams.Error.Count > 0)
                {
                    if (powershell.Streams.Error[0].FullyQualifiedErrorId == "NativeCommandError")
                    {
                        StringBuilder msgBuilder = new StringBuilder();
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
                // The error message is not an error, how annoying!
                string message = ret.First(x => (x.BaseObject as string).Length > 0).BaseObject as string;
                // Switch statement in C# only works on constants
                if (message == String.Format("Quota successfully created for \"{0}\".", path))
                {
                    successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                    return successMessage;
                }
                else if (message == String.Format("Quota already exists for \"{0}\".", path))
                {
                    // Check if the limit is the same
                    PowerShell powershell1 = PowerShell.Create();
                    PSCommand command1 = new PSCommand();
                    string script1 = String.Format("dirquota quota list /path:\"{0}\"", path);
                    command1.AddScript(script1);
                    powershell1.Commands = command1;
                    powershell1.Runspace = runspace;
                    Collection<PSObject> ret1 = powershell1.Invoke();
                    if (powershell1.Streams.Error.Count > 0)
                    {
                        if (powershell1.Streams.Error[0].FullyQualifiedErrorId == "NativeCommandError")
                        {
                            StringBuilder msgBuilder1 = new StringBuilder();
                            foreach(ErrorRecord errorRec1 in powershell1.Streams.Error)
                            {
                                // Kludge to fix a weird bug with blank lines in the error output
                                if (errorRec1.CategoryInfo.ToString() == errorRec1.Exception.Message)
                                {
                                    msgBuilder1.AppendLine();
                                }
                                else
                                {
                                    msgBuilder1.AppendLine(errorRec1.Exception.Message);
                                }
                            }
                            throw new Exception(msgBuilder1.ToString());
                        }
                        else
                        {
                            throw powershell1.Streams.Error[0].Exception;
                        }
                    }
                    string existingLimitLine = ret1.First(x => (x.BaseObject as String).StartsWith("Limit:")).BaseObject as string;
                    if (existingLimitLine == null)
                    {
                        // Must have been an error
                        string message1 = ret1.First(x => (x.BaseObject as string).Length > 0).BaseObject as string;
                        throw new Exception(message1);
                    }
                    else
                    {
                        // Parse out the quota and see if it is the same
                        // Separate the line into (Non-words)[bunch of spaces](Rest of line)"(Hard)"
                        GroupCollection groups = Regex.Match(existingLimitLine, @"(\S+)\s+(.+)\s\(.+\)$").Groups;
                        // The (Rest of line) part of the separation is the quota, expressed as e.g. "10.00 GB"
                        string existingLimit = groups[2].Value;
                        // Break out numeric and unit parts and compare them
                        GroupCollection limitGroups = Regex.Match(limit, @"(\d+)\s*(\D+)").Groups;
                        double limitNumber = double.Parse(limitGroups[1].Value);
                        string limitUnits = limitGroups[2].Value.ToUpper();
                        GroupCollection existingLimitGroups = Regex.Match(existingLimit, @"([.\d]+)\s*(\D+)").Groups;
                        double existingLimitNumber = double.Parse(existingLimitGroups[1].Value);
                        string existingLimitUnits = existingLimitGroups[2].Value.ToUpper();
                        if ((limitNumber != existingLimitNumber) || (limitUnits != existingLimitUnits))
                        {
                            throw new Exception(String.Format("Path '{0}' has limit '{1}', different than specified.", path, existingLimit));
                        }
                        else
                        {
                            successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                            return successMessage;
                        }
                    }
                }
                else
                {
                    throw new Exception(message);
                }

            }
            catch (Exception e)
            {
                errorMessage = new MSActorReturnMessageModel(ErrorCode, e.Message);
                return errorMessage;
            }

        }
    }
}