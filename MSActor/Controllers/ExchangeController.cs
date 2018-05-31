using MSActor.Models;
using Microsoft.Exchange.Data;
using Microsoft.Exchange.Data.Directory.Management;
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
    public class ExchangeController
    {
        public const string SuccessCode = "CMP";
        public const string ErrorCode = "ERR";
        public const string PendingCode = "PND";
        public const string RemoteExchangeScript = @". 'E:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'";
        UtilityController util;
        public ExchangeController()
        {
            util = new UtilityController();
        }
        public MSActorReturnMessageModel EnableMailbox(string database, string alias, string emailaddresses)
        {
            MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");

            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        runspace.Open();
                        powershell.Runspace = runspace;

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddCommand("Enable-Mailbox");
                        command.AddParameter("identity", alias);
                        command.AddParameter("database", database);
                        command.AddParameter("alias", alias);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            // Check if the mailbox exists and is the way we want it
                            using (PowerShell powershell1 = PowerShell.Create())
                            {
                                powershell1.Runspace = runspace;
                                command = new PSCommand();
                                command.AddCommand("Get-Mailbox");
                                command.AddParameter("Identity", alias);
                                powershell1.Commands = command;
                                Collection<PSObject> mailboxes = powershell1.Invoke();
                                if (powershell1.Streams.Error.Count > 0)
                                {
                                    // If the mailbox is not found, fall through and throw the other exception.
                                    // Otherwise something is probably really wrong and throw this exception instead.
                                    RemoteException ex1 = powershell1.Streams.Error[0].Exception as RemoteException;
                                    if (!ex1.SerializedRemoteException.TypeNames.Contains("Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException"))
                                    {
                                        throw powershell1.Streams.Error[0].Exception;
                                    }
                                }
                                Mailbox mailbox = mailboxes.FirstOrDefault()?.BaseObject as Mailbox;
                                if (mailbox != null
                                    && mailbox.Database.Name == database
                                    && mailbox.Alias == alias
                                    && mailbox.EmailAddresses.Contains(ProxyAddress.Parse("SMTP", emailaddresses))
                                   )
                                {
                                    return successMessage;
                                }
                                else
                                {
                                    throw powershell.Streams.Error[0].Exception;
                                }
                            }
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddCommand("set-mailbox");
                        command.AddParameter("identity", alias);
                        command.AddParameter("emailaddresses", emailaddresses);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
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

        public MSActorReturnMessageModel SetMailboxQuotas(string identity, string prohibitsendreceivequota, string prohibitsendquota, string issuewarningquota)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddCommand("Set-Mailbox");
                        command.AddParameter("Identity", identity);
                        command.AddParameter("IssueWarningQuota", issuewarningquota);
                        command.AddParameter("ProhibitSendQuota", prohibitsendquota);
                        command.AddParameter("ProhibitSendReceiveQuota", prohibitsendreceivequota);
                        command.AddParameter("UseDatabaseQuotaDefaults", false);
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

        public MSActorReturnMessageModel SetMailboxName(string identity, string alias, string addemailaddress)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Check that new alias does not already exist
                        command = new PSCommand();
                        command.AddCommand("Get-Mailbox");
                        command.AddParameter("Identity", alias);
                        powershell.Commands = command;
                        Collection<PSObject> existingMailboxes = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            // It's okay if the object is not found
                            RemoteException ex = powershell.Streams.Error[0].Exception as RemoteException;
                            if (!ex.SerializedRemoteException.TypeNames.Contains("Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException"))
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        powershell.Streams.ClearStreams();

                        if (existingMailboxes.Any(x => (x.BaseObject as Mailbox)?.Alias == alias))
                        {
                            throw new Exception("Mailbox for new alias already exists.");
                        }

                        command = new PSCommand();
                        command.AddCommand("Set-Mailbox");
                        command.AddParameter("Identity", identity);
                        command.AddParameter("Alias", alias);
                        if (addemailaddress != null)
                        {
                            Hashtable emailaddresses = new Hashtable
                            {
                                { "add", addemailaddress }
                            };
                            command.AddParameter("EmailAddresses", emailaddresses);
                        }
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

        public MSActorReturnMessageModel NewMoveRequest(string identity, string targetdatabase)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddCommand("Get-MoveRequest");
                        command.AddParameter("Identity", identity);
                        powershell.Commands = command;
                        Collection<PSObject> existingMoveRequests = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            RemoteException ex = powershell.Streams.Error[0].Exception as RemoteException;
                            // ManagementObjectNotFoundException is okay; it means there was not an existing move request
                            if (!ex.SerializedRemoteException.TypeNames.Contains("Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException"))
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        powershell.Streams.ClearStreams();

                        // If there already is a move request we need to figure out what to do about it
                        if (existingMoveRequests.Count > 0)
                        {
                            string moveRequestStatus = existingMoveRequests[0].Properties["Status"].Value.ToString();
                            if (moveRequestStatus != "Completed")
                            {
                                // Is the same move request in flight or are we conflicting with another one?
                                if (existingMoveRequests[0].Properties["TargetDatabase"].Value.ToString() == targetdatabase)
                                {
                                    MSActorReturnMessageModel pndMessage = new MSActorReturnMessageModel(PendingCode, "");
                                    return pndMessage;
                                }
                                else
                                {
                                    MSActorReturnMessageModel errMessage = new MSActorReturnMessageModel(ErrorCode, "Request still exists to move this mailbox to a different database");
                                    return errMessage;
                                }
                            }
                            else
                            // Remove the completed move request and go on to make a new one
                            {
                                command = new PSCommand();
                                command.AddCommand("Remove-MoveRequest");
                                command.AddParameter("Identity", identity);
                                command.AddParameter("Confirm", false);
                                powershell.Commands = command;
                                powershell.Invoke();
                                if (powershell.Streams.Error.Count > 0)
                                {
                                    throw powershell.Streams.Error[0].Exception;
                                }
                                powershell.Streams.ClearStreams();
                            }
                        }

                        command = new PSCommand();
                        command.AddCommand("New-MoveRequest");
                        command.AddParameter("Identity", identity);
                        command.AddParameter("TargetDatabase", targetdatabase);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            if (powershell.Streams.Error[0].Exception.Message.Contains("is already in the target database"))
                            {
                                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                                return successMessage;
                            }
                            else
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        else
                        {
                            MSActorReturnMessageModel pendingMessage = new MSActorReturnMessageModel(PendingCode, "");
                            return pendingMessage;
                        }
                        // powershell.Streams.ClearStreams();  -- is unreachable here
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }

        public MSActorReturnMessageModel GetMoveRequest(string identity)
        {
            // Multiple paths to error
            MSActorReturnMessageModel errorMessage;
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();

                        bool connected = false;
                        //if (powershell.Streams.Verbose.Contains(x => x.Message.StartsWith("Connected to")))
                        foreach (VerboseRecord vr in powershell.Streams.Verbose)
                        {
                            if (vr.Message.StartsWith("Connected to"))
                            {
                                connected = true;
                                break;
                            }
                        }

                        if (!connected && powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        command = new PSCommand();
                        command.AddCommand("Get-MoveRequest");
                        command.AddParameter("Identity", identity);
                        powershell.Commands = command;
                        Collection<PSObject> existingMoveRequests = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            // Here we are throwing an error on purpose if a move request does not exist
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        string status = existingMoveRequests[0].Properties["Status"].Value.ToString() as string;
                        switch (status)
                        {
                            case "Completed":
                                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");
                                return successMessage;
                            case "InProgress":
                            case "Queued":
                                MSActorReturnMessageModel pendingMessage = new MSActorReturnMessageModel(PendingCode, "");
                                return pendingMessage;
                            default:
                                string errorString = "Move request status is '" + status + "'";
                                errorMessage = new MSActorReturnMessageModel(ErrorCode, errorString);
                                Debug.WriteLine("ERROR: " + errorString);
                                return errorMessage;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                return util.ReportError(e);
            }
        }


        public MSActorReturnMessageModel DisableMailbox(string identity)
        {
            try
            {
                // For use later; there are multiple routes to success
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel(SuccessCode, "");

                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // First check for mobile devices and remove them
                        command = new PSCommand();
                        command.AddCommand("Get-MobileDevice");
                        command.AddParameter("Mailbox", identity);
                        powershell.Commands = command;
                        Collection<PSObject> mobileDevices = powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            // Mailbox may already be gone
                            RemoteException ex = powershell.Streams.Error[0].Exception as RemoteException;
                            if (ex.SerializedRemoteException.TypeNames.Contains("Microsoft.Exchange.Management.AirSync.RecipientNotFoundException"))
                            {
                                return successMessage;
                            }
                            else
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                        }
                        powershell.Streams.ClearStreams();

                        foreach (PSObject device in mobileDevices)
                        {
                            string deviceIdentity = device.Properties["Identity"].Value.ToString();
                            command = new PSCommand();
                            command.AddCommand("Remove-MobileDevice");
                            command.AddParameter("Identity", deviceIdentity);
                            command.AddParameter("Confirm", false);
                            powershell.Commands = command;
                            powershell.Invoke();
                            if (powershell.Streams.Error.Count > 0)
                            {
                                throw powershell.Streams.Error[0].Exception;
                            }
                            powershell.Streams.ClearStreams();
                        }

                        command = new PSCommand();
                        command.AddCommand("Disable-Mailbox");
                        command.AddParameter("Identity", identity);
                        command.AddParameter("Confirm", false);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
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

        public MSActorReturnMessageModel HideMailboxFromAddressLists(string identity, string hidemailbox)
        {
            try
            {
                PSSessionOption option = new PSSessionOption();
                using (PowerShell powershell = PowerShell.Create())
                {
                    using (Runspace runspace = RunspaceFactory.CreateRunspace())
                    {
                        powershell.Runspace = runspace;
                        runspace.Open();

                        PSCommand command = new PSCommand();
                        // We get errors when the Exchange remote script tries to talk to us,
                        // unless we redefine Write-Host to be an empty script.
                        command.AddScript("Function Write-Host {}");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Load the Exchange Management Shell startup script
                        command = new PSCommand();
                        command.AddScript(RemoteExchangeScript);
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Ask Exchange Management Shell to sniff the environment for a server and give us a session
                        command = new PSCommand();
                        command.AddCommand("Connect-ExchangeServer");
                        command.AddParameter("Auto");
                        powershell.Commands = command;
                        powershell.Invoke();
                        if (powershell.Streams.Error.Count > 0)
                        {
                            throw powershell.Streams.Error[0].Exception;
                        }
                        powershell.Streams.ClearStreams();

                        // Now set the HiddenFromAddressListsEnabled flag
                        command = new PSCommand();
                        command.AddCommand("Set-Mailbox");
                        command.AddParameter("Identity", identity);
                        command.AddParameter("HiddenFromAddressListsEnabled", Boolean.Parse(hidemailbox));
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
    }

}