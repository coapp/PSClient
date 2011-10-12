using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace CoApp.PSClient.Remoting
{

    public static class Remote
    {
        /*
         * It should be noted that at the time of this writing, there is effectively no existing public
         *   documentation for the following remote connection methods from Microsoft.  These sections
         *   were completed based upon documentation produced by the trial and error of others.
         *   
         * Base source for connections:
         *   http://com2kid.wordpress.com/2011/09/22/remotely-executing-commands-in-powershell-using-c/
         * 
         */
        public class ConnectionGenerator
        {
            public static readonly int PortSSL = 5986;
            public static readonly int PortNoSSL = 5985;
            public static readonly string DefaultAppName = "/wsman";
            public static readonly string DefaultShellUri = "http://schemas.microsoft.com/powershell/Microsoft.PowerShell";
            private readonly bool _UseSSL;
            private readonly PSCredential _Credential;
            private readonly int _Port;

            public ConnectionGenerator(bool UseSSL, PSCredential Credential, int? Port = null)
            {
                _UseSSL = UseSSL;
                _Credential = Credential;
                _Port = Port ?? (UseSSL ? PortSSL : PortNoSSL);
            }

            public Runspace Generate(string RemoteComputer, int? timeout = null, bool CheckCLR = false)
            {
                WSManConnectionInfo info;
                if (timeout != null)
                    info = new WSManConnectionInfo(_UseSSL, RemoteComputer, _Port, DefaultAppName, DefaultShellUri, _Credential, timeout.Value);
                else
                    info = new WSManConnectionInfo(_UseSSL, RemoteComputer, _Port, DefaultAppName, DefaultShellUri, _Credential);
                Runspace runspace = RunspaceFactory.CreateRunspace(info);

                runspace.Open();
                if (CheckCLR)
                {
                    Pipeline pipe = runspace.CreatePipeline();
                    pipe.Commands.AddScript("$PSVersionTable.Item('CLRVersion').Major");
                    Collection<PSObject> res = pipe.Invoke();
                    pipe.Dispose();
                    pipe = runspace.CreatePipeline();
                    if (res != null)
                    {
                        int ver = Int32.Parse(res.FirstOrDefault().ToString());
                        if (ver < 4)
                        {
                            // need to write a system file to push the 4.0 CLR
                            string pathScript = @"$f = (dir env:SystemRoot).value + '\System32\wsmprovhost.exe.config'";
                            pipe.Commands.AddScript(pathScript);
                            pipe.Invoke();
                            pipe.Dispose();
                            pipe = runspace.CreatePipeline();
                            string contentScript = @"'<?xml version=""1.0"" encoding=""utf-8"" ?>
<configuration>
  <startup useLegacyV2RuntimeActivationPolicy=""true"">
    <supportedRuntime version=""v4.0""/>
  </startup>
</configuration>'";
                            string commandScript = @"Set-Content -Path $f -Force -Encoding UTF8 -Value ";
                            pipe.Commands.AddScript(commandScript + contentScript);
                            pipe.Invoke();
                        }
                    }
                    else // bad things happened
                        throw new RuntimeException("Error receiving objects from remote connection.");

                    // Clean up and make a new runspace to hand off.
                    try
                    {
                        pipe.Dispose();
                        runspace.Close();
                    }
                    catch (Exception)
                    { }
                    runspace = RunspaceFactory.CreateRunspace(info);
                    runspace.Open();
                }

                return runspace;
            }
        }

        public static Collection<PSObject> Exec(Runspace runspace, Command command)
        {
            PowerShell shell = PowerShell.Create();
            shell.Runspace = runspace;

            Command import = new Command("Import-Module");
            import.Parameters.Add("Name", "CoApp");
            shell.Commands.AddCommand(import);
            shell.Invoke();
            shell.Commands.Clear();
            shell.Commands.AddScript(command.ToString());
            Collection<PSObject> ret = shell.Invoke();
            try
            {
                shell.Dispose();
                runspace.Close();
            }
            catch (Exception)
            { }
            return ret;
        }

        public static Collection<PSObject> Exec(Runspace runspace, string script)
        {
            runspace.Open();
            PowerShell shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript(script);
            Collection<PSObject> ret = shell.Invoke();
            try
            {
                shell.Dispose();
                runspace.Close();
            }
            catch (Exception)
            { }
            return ret;
        }

        internal class RemoteOutput
        {
            public string RemoteName;
            public List<PSObject> Output;
            public RemoteOutput(string Name)
            {
                RemoteName = Name;
                Output = new List<PSObject>();
            }
        }

        [Cmdlet(VerbsLifecycle.Invoke, "RemoteCommand")]
        public class Invoke_RemoteCommand : PSCmdlet
        {
            [Parameter(Mandatory = true, ValueFromPipeline = false)]
            public string[] ComputerName;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public PSCredential Credential = null;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public SwitchParameter UseSSL = false;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public int? Port = null;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public int? Timeout = null;
            [Parameter(Mandatory = true, ValueFromPipeline = true)]
            public string Command;

            protected override void ProcessRecord()
            {
                ConnectionGenerator generator = new ConnectionGenerator(UseSSL, Credential, Port);
                foreach (string comp in ComputerName)
                {
                    Collection<PSObject> result = Exec(generator.Generate(comp, Timeout), this.Command);
                    RemoteOutput O = new RemoteOutput(comp);
                    if (result != null)
                        O.Output.AddRange(result);
                    else
                        O.Output = null;

                    WriteObject(O);
                }
            }
        }

        [Cmdlet(VerbsLifecycle.Start, "RemoteSession")]
        public class Start_RemoteSession : PSCmdlet
        {
            [Parameter(Mandatory = true, ValueFromPipeline = false)]
            public string ComputerName;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public PSCredential Credential = null;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public SwitchParameter UseSSL = false;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public int? Port = null;
            [Parameter(Mandatory = false, ValueFromPipeline = false)]
            public int? Timeout = null;

            protected override void ProcessRecord()
            {
                ConnectionGenerator generator = new ConnectionGenerator(UseSSL, Credential, Port);
                Runspace R = generator.Generate(ComputerName, Timeout);
                R.Open();
                PowerShell shell = PowerShell.Create();
                shell.Runspace = R;
                string script = Host.UI.ReadLine();
                while (!script.Equals("exit", StringComparison.CurrentCultureIgnoreCase))
                {
                    shell.Commands.Clear();
                    shell.AddScript(script);
                    Collection<PSObject> ret = shell.Invoke();
                    if (ret != null)
                        WriteObject(ret, true);
                    Host.UI.WriteLine();
                    Host.UI.Write(ComputerName + "> ");
                    script = Host.UI.ReadLine();
                }

                try
                {
                    shell.Dispose();
                    R.Close();
                }
                catch (Exception)
                { }
            }
        }
    }
}
