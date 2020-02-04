using Microsoft.Management.Infrastructure;
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
        public const long BytesInGigabyte = 1_073_741_824;

        UtilityController util;
        public FileServerController()
        {
            util = new UtilityController();
        }

        public MSActorReturnMessageModel NewDirectory(string computername, string path)
        {
            try
            {
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");

                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        command.AddCommand("New-Item");
                        command.AddParameter("ItemType", "directory");
                        command.AddParameter("Path", path);
                        powershell.Commands = command;
                        Collection<PSObject> returns = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            if (powershell.Streams.Error[0].Exception.Message == String.Format("Item with specified name {0} already exists.", path))
                            {
                                return successMessage;
                            }
                            else
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        powershell.Streams.ClearStreams();

                        return successMessage;
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        public MSActorReturnMessageModel RemoveDirectory(string computername, string path)
        {
            try
            {
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");

                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        command.AddCommand("Remove-Item");
                        command.AddParameter("Path", path);
                        command.AddParameter("Recurse");
                        command.AddParameter("Force");
                        powershell.Commands = command;
                        powershell.Invoke();

                        if (powershell.Streams.Error.Count > 0)
                        {
                            RemoteException ex = powershell.Streams.Error[0].Exception as RemoteException;
                            if (ex.SerializedRemoteException.TypeNames.Contains("Deserialized.System.Management.Automation.ItemNotFoundException"))
                            {
                                return successMessage;
                            }
                            else if (ex.SerializedRemoteException.TypeNames.Contains("Deserialized.System.IO.PathTooLongException"))
                            {
                                // Run our script for extra long paths instead
                                using (PowerShell powershell1 = PowerShell.Create())
                                {
                                    powershell1.Runspace = runspace;

                                    PSCommand command1 = new PSCommand();
                                    command1.AddCommand("Set-ExecutionPolicy");
                                    command1.AddParameter("ExecutionPolicy", "RemoteSigned");
                                    command1.AddParameter("Scope", "Process");
                                    command1.AddParameter("Force");
                                    powershell1.Commands = command1;
                                    powershell1.Invoke();
                                    if (powershell1.Streams.Error.Count > 0)
                                    {
                                        throw powershell1.Streams.Error[0].Exception;
                                    }
                                    powershell1.Streams.ClearStreams();

                                    command1 = new PSCommand();
                                    command1.AddScript(". D:\\PathTooLong.ps1");
                                    powershell1.Commands = command1;
                                    powershell1.Invoke();
                                    if (powershell1.Streams.Error.Count > 0)
                                    {
                                        throw powershell1.Streams.Error[0].Exception;
                                    }
                                    powershell1.Streams.ClearStreams();

                                    command1 = new PSCommand();
                                    command1.AddCommand("Remove-PathToLongDirectory");
                                    command1.AddArgument(path);
                                    powershell1.Commands = command1;
                                    powershell1.Invoke();
                                    if (powershell1.Streams.Error.Count > 0)
                                    {
                                        throw powershell1.Streams.Error[0].Exception;
                                    }
                                    powershell1.Streams.ClearStreams();
                                }
                            }
                            else
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        powershell.Streams.ClearStreams();

                        return successMessage;
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        public MSActorReturnMessageModel RenameDirectory(string computername, string path, string newname)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        command.AddCommand("Rename-Item");
                        command.AddParameter("Path", path);
                        command.AddParameter("NewName", newname);
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
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }
        public MSActorReturnMessageModel AddNetShare(string name, string computername, string path)
        {
            MSActorReturnMessageModel successMessage;

            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        string script = String.Format("net share {0}={1} \"/GRANT:Everyone,Full\"", name, path);
                        command.AddScript(script);
                        powershell.Commands = command;
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
                                    powershell.Streams.ClearStreams();

                                    // Check that the existing share has the same path
                                    command = new PSCommand();
                                    script = String.Format("net share {0}", name);
                                    command.AddScript(script);
                                    powershell.Commands = command;
                                    Collection<PSObject> ret = powershell.Invoke();
                                    if (powershell.Streams.Error.Count > 0)
                                    {
                                        throw powershell.Streams.Error[0].Exception;
                                    }
                                    powershell.Streams.ClearStreams();

                                    // Find the first (hopefully only) line in the output with "Path" at the beginning
                                    string pathResultLine = ret.First(x => (x.BaseObject as string).StartsWith("Path")).BaseObject as string;
                                    if (pathResultLine == null)
                                    {
                                        // There was not a line in the output containing the path, so we assume we got an error message instead.
                                        string message = ret.First(x => (x.BaseObject as string).Length > 0).BaseObject as string;
                                        throw new Exception(message);
                                    }
                                    else
                                    {
                                        // The output looks like "Path              D:\Users\srenker".
                                        // The regular expression below separates this into groups.
                                        // Meaning of the next regex (from left to right):
                                        // 1. Save all the characters that are not blanks into a group. (\S+)
                                        // 2. Skip over all characters that are blanks. \s+
                                        // 3. Save all the other characters into a group, up to end of line. (.+)$
                                        // It's done this way because the path may have a space embedded in the name.
                                        // The @ before the string tells C# not to escape any characters before passing it
                                        // to the regular expression processor.
                                        GroupCollection groups = Regex.Match(pathResultLine, @"(\S+)\s+(.+)$").Groups;
                                        // Group 2 (#3 above) is the path value.
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
                            }
                            else
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                            // powershell.Streams.ClearStreams();  -- is unreachable here
                        }

                        successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                        return successMessage;
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        public MSActorReturnMessageModel RemoveNetShare(string name, string computername, string path)
        {
            MSActorReturnMessageModel successMessage;

            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        // First check that the share name is for the correct path
                        PSCommand command = new PSCommand();
                        string script = String.Format("net share {0}", name);
                        command.AddScript(script);
                        powershell.Commands = command;
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
                        powershell.Streams.ClearStreams();

                        // Find the first (hopefully only) line in the output with "Path" at the beginning
                        string pathResultLine = ret.First(x => (x.BaseObject as string).StartsWith("Path")).BaseObject as string;
                        // The output looks like "Path              D:\Users\srenker".
                        // The regular expression below separates this into groups.
                        // Meaning of the next regex (from left to right):
                        // 1. Save all the characters that are not blanks into a group. (\S+)
                        // 2. Skip over all characters that are blanks. \s+
                        // 3. Save all the other characters into a group, up to end of line. (.+)$
                        // It's done this way because the path may have a space embedded in the name.
                        // The @ before the string tells C# not to escape any characters before passing it
                        // to the regular expression processor.
                        GroupCollection groups = Regex.Match(pathResultLine, @"(\S+)\s+(.+)$").Groups;
                        // Group 2 (#3 above) is the path value.
                        string existingPath = groups[2].Value;
                        if (!String.Equals(path, existingPath, StringComparison.OrdinalIgnoreCase))
                        {
                            throw new Exception(String.Format("Share '{0}' is for path '{1}', different than specified.", name, existingPath));
                        }

                        // Now delete the share
                        script = String.Format("net share {0} /delete", name);
                        command.AddScript(script);
                        powershell.Commands = command;
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
                        powershell.Streams.ClearStreams();

                        successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                        return successMessage;
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        public MSActorReturnMessageModel AddUserFolderAccess(string employeeid, string samaccountname, string computername, string path, string accesstype)
        {
            try
            {
                PSObject user = util.getADUser(employeeid, samaccountname);
                if (user == null)
                {
                    MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel(ErrorCode, "User was not found.");
                    var customEx = new Exception("User was not found", new Exception());
                    Elmah.ErrorSignal.FromCurrentContext().Raise(customEx);
                    return errorMessage;
                }
                else
                {
                    string identity = user.Properties["SamAccountName"].Value as string;
                    return AddFolderAccess(identity, computername, path, accesstype);
                }
            }catch(Exception e)
            {
                return util.ReportError(e);
            }
        }

        public MSActorReturnMessageModel AddGroupFolderAccess(string groupname, string computername, string path, string accesstype)
        {
            return AddFolderAccess(groupname, computername, path, accesstype);
        }

        private MSActorReturnMessageModel AddFolderAccess(string identity, string computername, string path, string accesstype)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        // Note: The commands stacked on top of each other (prior to invoking) below have the effect
                        // of piping the output of one into the other, e.g. the result of Get-Acl becomes an input to Set-Variable.
                        // We need to work this way on a remote session so that type information does not get changed by
                        // retrieving the objects back to the local session.

                        PSCommand command = new PSCommand();
                        command.AddCommand("Get-Acl");
                        command.AddParameter("Path", path);
                        command.AddCommand("Set-Variable");
                        command.AddParameter("Name", "acl");
                        powershell.Commands = command;
                        Collection<PSObject> result = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

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
                        result = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddScript("$acl.SetAccessRule($perms)");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddScript(String.Format("Set-Acl -AclObject $acl -Path {0}", path));
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
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        private ulong NumericLimit(string limit)
        {
            int suffixIndex = limit.IndexOf("gb", StringComparison.InvariantCultureIgnoreCase);
            if (suffixIndex == -1)
                throw new Exception("Cannot process limit value - gigabytes only");
            return ulong.Parse(limit.Substring(0, suffixIndex).Trim()) * BytesInGigabyte;
        }

        public MSActorReturnMessageModel AddDirQuota(string computername, string path, string limit)
        {
            // Project P0975: Replace old command line scripts with new PowerShell commands,
            // required after upgrading the Windows Server version on the file servers
            MSActorReturnMessageModel successMessage;
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        command.AddCommand("New-FsrmQuota");
                        command.AddParameter("Path", path);
                        command.AddParameter("Size", NumericLimit(limit));
                        powershell.Commands = command;
                        Collection<PSObject> result = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            if (powershell.Streams.Error[0].Exception.Message.Trim() != "0x80045303, The specified object already exists.")
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                            else
                            {
                                powershell.Streams.ClearStreams();

                                // Check that the existing quota has the same limit
                                command = new PSCommand();
                                command.AddCommand("Get-FsrmQuota");
                                command.AddParameter("Path", path);
                                powershell.Commands = command;
                                Collection<PSObject> res = powershell.Invoke();
                                if (powershell.Streams.Error.Count > 0)
                                {
                                    throw powershell.Streams.Error[0].Exception;
                                }
                                CimInstance quota = (CimInstance)res.FirstOrDefault()?.BaseObject;
                                if ((ulong)quota.CimInstanceProperties["Size"].Value != NumericLimit(limit))
                                {
                                    throw new Exception("A different quota already exists on that folder");
                                }
                            }
                        }
                        powershell.Streams.ClearStreams();

                        successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                        return successMessage;
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }

        }

        public MSActorReturnMessageModel ModifyDirQuota(string computername, string path, string limit)
        {
            // Project P0975: Replace old command line scripts with new PowerShell commands,
            // required after upgrading the Windows Server version on the file servers
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    string url = String.Format("http://{0}:5985/wsman", computername);
                    Uri uri = new Uri(url);
                    WSManConnectionInfo conn = new WSManConnectionInfo(uri);
                    using (Runspace runspace = RunspaceFactory.CreateRunspace(conn))
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        command.AddCommand("Set-FsrmQuota");
                        command.AddParameter("Path", path);
                        command.AddParameter("Size", NumericLimit(limit));
                        powershell.Commands = command;
                        Collection<PSObject> result = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                        return successMessage;
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }
    }
}