using MSActor.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Web;

namespace MSActor.Controllers
{
    public class ExchangeController
    {
        public ExchangeController()
        {

        }

        public MSActorReturnMessageModel EnableUserMailboxDriver(string identity, string database, string alias)
        {
            try
            {
                PowerShell ps = PowerShell.Create();
                ps.AddCommand("Add-PSSnapin");
                ps.AddParameter("name", "Microsoft.Exchange.Management.PowerShell.SnapIn");
                ps.Invoke();
                ps = PowerShell.Create();
                ps.AddCommand("Enable-Mailbox");
                ps.AddParameter("Identity", identity);
                ps.AddParameter("database", database);
                ps.AddParameter("alias", alias);
                ps.Invoke();
                MSActorReturnMessageModel successMessage = new MSActorReturnMessageModel("SUCCESS", "");
                return successMessage;
            }
            catch(Exception e)
            {
                MSActorReturnMessageModel errorMessage = new MSActorReturnMessageModel("ERROR", e.Message);
                Debug.WriteLine("ERROR: " + e.Message);
                return errorMessage;
            }
        }
    }
}