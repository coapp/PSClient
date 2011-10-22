using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.DirectoryServices;

namespace PSClient
{
    static class DomainHelper
    {
        static internal IEnumerable<string> ListComputers(string domain = null, string user = null, string pass = null)
        {
            DirectoryEntry DE;
            if (domain == null)
                DE = new DirectoryEntry();
            else
                if (user != null && pass != null)
                    DE = new DirectoryEntry(domain, user, pass);
                else
                    DE = new DirectoryEntry(domain);
            
            DirectorySearcher DS = new DirectorySearcher(DE, "(objectCategory=computer)", new []{"name"});
            SearchResultCollection result = DS.FindAll();
            List<string> values = (from SearchResult item in result select item.Properties into obj select obj["name"].ToString()).ToList();
            return values;
        }
    }

    [Cmdlet(VerbsCommon.Get, "DomainComputers")]
    public class Get_DomainComputers : PSCmdlet
    {
        [Parameter(Mandatory = false, Position = 0)]
        public string DomainName;
        [Parameter(Mandatory = false)]
        public PSCredential Credential;

        protected override void ProcessRecord()
        {
            WriteObject(DomainHelper.ListComputers(DomainName, Credential.UserName, Credential.GetNetworkCredential().Password));
            
        }
    }
}
