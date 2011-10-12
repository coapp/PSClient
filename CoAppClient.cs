//-----------------------------------------------------------------------
// <copyright company="CoApp Project">
//     Copyright (c) 2011 Timothy Rogers. All rights reserved.
// </copyright>
// <license>
//     The software is licensed under the Apache 2.0 License (the "License")
//     You may not use the software except in compliance with the License. 
// </license>
//-----------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Net;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using CoApp.PSClient.Remoting;
using CoApp.Toolkit.Engine;
using CoApp.Toolkit.Extensions;
using CoApp.Toolkit.Engine.Client;

namespace CoApp.PSClient
{
    #region SnapIn (Old)
    /* Old section.  Used when this was a snapin instead of a module.
     * 
     * Powershell SnapIn client for CoApp
     * 
     * Note:  To register the snapin with the system (and therefore use a single 
     *        command to load the cmdlets), enter the following at the powershell 
     *        prompt from the library directory.
     * 
     *    Invoke-Expression ((Get-ChildItem -Recurse -Path "${env:systemroot}\Microsoft.NET" -include installutil.exe)[-1].ToString() + " .\CoApp.dll")
     *    
     *        After doing this once, the cmdlets may be loaded in the future by
     *        entering the following (not case-sensitive).
     *        
     *    Add-PSSnapin CoApp
     *    
     *        Also, be aware that these cmdlets require Powershell to use the .NET 4 runtimes, which is not loaded by default in Powershell 2.0.
     */ 
     /*
    [RunInstaller(true)]
    public class PSClient_Snapin : PSSnapIn
    {
        public override string Name
        { get { return "CoApp"; } }

        public override string Description
        { get { return "PowerShell cmdlets for CoApp."; } }

        public override string Vendor
        { get { return "CoApp.org"; } }

        public override string[] Formats
        {
            get { return new string[] { "CoApp.format.ps1xml" }; }
        }
    }
    */
    #endregion

    internal class Feed
    {
        public String Location;
        public DateTime LastScanned;
        public bool Session;
        public bool Supressed;
        public bool Verified;
    }

    public class CoApp_Cmdlet : PSCmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public string SessionID = null;

        // remoting parameters
        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public string[] ComputerName = null;
        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public PSCredential Credential = null;
        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public SwitchParameter UseSSL = false;
        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public int? Port = null;
        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public int? Timeout = null;

        protected List<Runspace> Remotes;
        
        protected PackageManager PM;
        protected PackageManagerMessages messages;

        // Internal dictionary.  Is only used for selection prompts.
        internal static Toolkit.Collections.EasyDictionary<int, string> Selections = new Toolkit.Collections.EasyDictionary<int,string>
            {
                {0, "&0 - "},
                {1, "&1 - "},
                {2, "&2 - "},
                {3, "&3 - "},
                {4, "&4 - "},
                {5, "&5 - "},
                {6, "&6 - "},
                {7, "&7 - "},
                {8, "&8 - "},
                {9, "&9 - "},
                {10, "&a - "},
                {11, "&b - "},
                {12, "&c - "},
                {13, "&d - "},
                {14, "&e - "},
                {15, "&f - "},
                {16, "&g - "},
                {17, "&h - "},
                {18, "&i - "},
                {19, "&j - "},
                {20, "&k - "},
                {21, "&l - "},
                {22, "&m - "},
                {23, "&n - "},
                {24, "&o - "},
                {25, "&p - "},
                {26, "&q - "},
                {27, "&r - "},
                {28, "&s - "},
                {29, "&t - "},
                {30, "&u - "},
                {31, "&v - "},
                {32, "&w - "},
                {33, "&x - "},
                {34, "&y - "},
                {35, "&z - "},
            };
        protected List<Object> output;
        protected Task task;
        protected Cmdlet BackRef;
            
        protected override void BeginProcessing()
        {
            BackRef = this;
            output = new List<object>();

            // Remoting...
            if (!ComputerName.IsNullOrEmpty())
            {
                Remotes = new List<Runspace>();
                Remote.ConnectionGenerator Gen = new Remote.ConnectionGenerator(UseSSL, Credential, Port);
                foreach (string comp in ComputerName)
                {
                    Remotes.Add(Gen.Generate(comp, Timeout, true));
                }
            }
            // Not remoting...
            else
            {
                PM = PackageManager.Instance;
                PM.Connect("PSClient", SessionID);
                messages = new PackageManagerMessages
                               {
                                   UnexpectedFailure = UnexpectedFailure,
                                   NoPackagesFound = NoPackagesFound,
                                   PermissionRequired = OperationRequiresPermission,
                                   Error = MessageArgumentError,
                                   RequireRemoteFile = GetRemoteFile,
                                   OperationCancelled = CancellationRequested,
                                   PackageInformation = PackageInfo,
                                   PackageDetails = PackageDetail,
                                   FeedDetails = FeedInfo,
                                   ScanningPackagesProgress = ScanProgress,
                                   InstallingPackageProgress = InstallProgress,
                                   RemovingPackageProgress = RemoveProgress,
                                   InstalledPackage = InstallComplete,
                                   RemovedPackage = RemoveComplete,
                                   FailedPackageInstall = InstallFailed,
                                   FailedPackageRemoval = RemoveFailed,
                                   Warning = WarningMsg,
                                   FeedAdded = NewFeed,
                                   FeedRemoved = LostFeed,
                                   FeedSuppressed = SupressedFeed,
                                   NoFeedsFound = NoFeeds,
                                   FileNotFound = NoFile,
                                   PackageBlocked = BlockedPackage,
                                   FileNotRecognized = NotRecognized,
                                   Recognized = Recognized,
                                   UnknownPackage = Unknown,
                                   PackageHasPotentialUpgrades = PotentialUpgrades,
                                   UnableToDownloadPackage = UnableToDownload,
                                   UnableToInstallPackage = UnableToInstall,
                                   UnableToResolveDependencies = UnableToResolveDeps,
                                   PackageSatisfiedBy = Satisfied
                               };
            }
            
        }

        protected override void EndProcessing()
        {
            WaitForComplete();
            if (ComputerName.IsNullOrEmpty())
                PM.Disconnect();
        }

        #region Default message handlers
        protected void PackageInfo(Package P)
            { Host.UI.WriteLine("PackageInfo received: " + P.CanonicalName); }
        protected void PackageDetail(Package P)
            { Host.UI.WriteLine("PackageDetail received: " + P.CanonicalName); }
        protected void FeedInfo(string Location, DateTime Scanned, bool Session, bool Supressed, bool Validated)
            { Host.UI.WriteLine("FeedInfo received: " + Location); }
        protected void ScanProgress(string Location, int Progress)
            { Host.UI.WriteLine("Scanning in progress:  " + Location + "  " + Progress + "%"); }
        protected void InstallProgress(string CName, int Progress, int TotalProgress)
        { Host.UI.WriteLine("Install in progress:  " + CName + "  " + Progress + "%" + "  Total: " + TotalProgress + "%"); }
        protected void RemoveProgress(string CName, int Progress)
            { Host.UI.WriteLine("Remove in progress:  " + CName + "  " + Progress + "%"); }
        protected void InstallComplete(string CName)
            { Host.UI.WriteLine("Install complete:  " + CName); }
        protected void RemoveComplete(string CName)
            { Host.UI.WriteLine("Remove complete:  " + CName); }
        protected void InstallFailed(string CName, string Filename, string Reason)
            { Host.UI.WriteLine("Install failed for " + CName + " from " + Filename + ":  " + Reason); }
        protected void RemoveFailed(string CName, string Reason)
            { Host.UI.WriteLine("Remove failed for " + CName + ":  " + Reason); }
        protected void WarningMsg(string arg1, string arg2, string arg3)
            { Host.UI.WriteLine("Warning:  " + arg1 + ", " + arg2 + ", " + arg3); }
        protected void NewFeed(string feed)
            { Host.UI.WriteLine("Feed added:  " + feed); }
        protected void LostFeed(string feed)
            { Host.UI.WriteLine("Feed removed:  " + feed); }
        protected void SupressedFeed(string feed)
            { Host.UI.WriteLine("Feed supressed:  " + feed); }
        protected void NoFeeds()
            { Host.UI.WriteLine("No feeds found."); }
        protected void NoFile(string Filename)
            { Host.UI.WriteLine("File not found:  " + Filename); }
        protected void BlockedPackage(string CName)
            { Host.UI.WriteLine("Package is blocked:  " + CName); }
        protected void NotRecognized(string Location, string Remote)
            { Host.UI.WriteLine("File not recognized:  " + Location + " from " + Remote); }
        protected void Recognized(string CName)
            { Host.UI.WriteLine("File recognized:  " + CName); }
        protected void Unknown(string CName)
            { Host.UI.WriteLine("Unknown package:  " + CName); }
        protected void PotentialUpgrades(Package Current, IEnumerable<Package> Super)
            { Host.UI.WriteLine("Package may have upgrades available:  " + Current.CanonicalName); }
        protected void UnableToDownload(Package Pack)
            { Host.UI.WriteLine("Unable to download package:  " + Pack.CanonicalName); }
        protected void UnableToInstall(Package Pack)
            { Host.UI.WriteLine("Unable to install package:  " + Pack.CanonicalName); }
        protected void UnableToResolveDeps(Package Pack, IEnumerable<Package> Deps)
            { Host.UI.WriteLine("Unable to resolve package dependencies:  " + Pack.CanonicalName); }
        protected void Satisfied(Package Requested, Package Existing)
            { Host.UI.WriteLine("Package requirement is satisfied:  " + Requested.CanonicalName + " ==> " + Existing.CanonicalName); }
        protected void CancellationRequested(string obj)
            { Host.UI.WriteLine("Operation Canceled:  " + obj); }
        protected void MessageArgumentError(string arg1, string arg2, string arg3)
            { Host.UI.WriteLine("Message Argument Error " + arg1 + ", " + arg2 + ", " + arg3 + "."); }
        protected void OperationRequiresPermission(string policyName)
            { Host.UI.WriteLine("Operation requires permission Policy: " + policyName); }
        protected void NoPackagesFound()
            { Host.UI.WriteLine("Did not find any packages."); }
        protected void UnexpectedFailure(Exception obj)
            { Host.UI.WriteLine("SERVER EXCEPTION: " + obj.Message + "\n" + obj.StackTrace); }
        #endregion

        protected void GetRemoteFile(string canonicalName, IEnumerable<string> arg2, string arg3, bool arg4)
        {
            string remote = String.Empty;
            bool getFile = true;
            if (!arg4)
                if (File.Exists(Path.Combine(arg3, canonicalName)))
                    getFile = false;
            if (getFile)
            {
                int i = 0;
                bool running = true;
                WebClient WC = new WebClient();
                WC.DownloadFileCompleted += new AsyncCompletedEventHandler((O, Args) =>
                {
                    if (File.Exists(Path.Combine(arg3, canonicalName)))
                        PM.RecognizeFile(canonicalName, Path.Combine(arg3, canonicalName), remote, messages);
                    else
                    {
                        // didn't really grab the file, move on to next source
                        running = true;
                    }
                });
                IEnumerator<string> R = arg2.GetEnumerator();
                while (!(File.Exists(Path.Combine(arg3, canonicalName))) && i++ < arg2.Count() && running)
                {
                    R.MoveNext();
                    remote = R.Current;
                    WC.DownloadFileAsync(new Uri(remote), Path.Combine(arg3, canonicalName));
                    running = false;
                }
                if (!File.Exists(Path.Combine(arg3, canonicalName)))
                {
                    PM.UnableToAcquire(canonicalName,messages);
                }
            }
            if (!getFile)
                PM.RecognizeFile(canonicalName, Path.Combine(arg3, canonicalName), null, messages);
            /*
             * To be used with the TransferManager when it's done
            if (File.Exists(Path.Combine(arg3, canonicalName)))
                PM.RecognizeFile(canonicalName, Path.Combine(arg3, canonicalName), remote);
            else
                PM.UnableToAcquire(canonicalName);
             */
        }

        private void WaitForComplete()
        {
            if (task != null && !(task.IsCompleted || task.IsCanceled))
                try
                {
                    task.Wait();
                }
                catch(Exception e)
                {
                    Host.UI.WriteLine(ConsoleColor.Red, Host.UI.RawUI.BackgroundColor, "Task faulted.");
                }
            // wait for cancellation token, or service to disconnect
            if (ComputerName.IsNullOrEmpty())
                WaitHandle.WaitAny(new[] {PM.IsDisconnected, PM.IsCompleted});
            
            


            foreach (var item in output)
            {
                WriteObject(item);
                FormatTable table = new FormatTable(new string[] {});
            }
        }

    }

    [Cmdlet("List", "Package", DefaultParameterSetName = "Typed")]
    public class ListPackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = false, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = false, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = false, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false)]
        public string MinVersion;
        [Parameter(Mandatory = false)]
        public string MaxVersion;
        [Parameter(Mandatory = false)]
        public bool? Dependencies;
        [Parameter(Mandatory = false)]
        public bool? Installed;
        [Parameter(Mandatory = false)]
        public bool? Active;
        [Parameter(Mandatory = false)]
        public bool? Required;
        [Parameter(Mandatory = false)]
        public bool? Blocked;
        [Parameter(Mandatory = false)]
        public bool? Latest;
        [Parameter(Mandatory = false)]
        public int? Index;
        [Parameter(Mandatory = false)]
        public int? MaxResults;
        [Parameter(Mandatory = false)]
        public string Location;
        [Parameter(Mandatory = false)]
        public SwitchParameter ForceScan;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("List-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Installed != null) com.Parameters.Add("Installed", Installed);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                                    {
                                        foreach (Runspace R in Remotes)
                                        {
                                            Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() => 
                                                Remote.Exec(R, com));
                                            T.ContinueWith(antecedent => 
                                                    output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                                            T.Start();
                                        }
                                    });
                task.Start();
            }
            //Non-Remoting
            else
            {
                switch (ParameterSetName)
                {
                    case "Package":
                        task = PM.GetPackages(InputPackage.CanonicalName, null, null, null, null, null, null,
                                              null, null, null, null, messages).ContinueWith(FP =>
                                                                                                 {
                                                                                                     foreach (
                                                                                                         Package P in
                                                                                                             FP.Result)
                                                                                                         output.Add(P);
                                                                                                 },
                                                                                             TaskContinuationOptions.
                                                                                                 AttachedToParent);
                        break;
                    case "Canonical":
                        task = PM.GetPackages(CanonicalName, MinVersion.VersionStringToUInt64(),
                                              MaxVersion.VersionStringToUInt64(), Dependencies,
                                              Installed,
                                              Active, Required, Blocked, Latest, Location,
                                              ForceScan,
                                              messages).ContinueWith((FP) =>
                                                                         {
                                                                             foreach (Package P in FP.Result)
                                                                                 output.Add(P);
                                                                         }, TaskContinuationOptions.AttachedToParent);
                        break;
                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") +
                                        (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(
                            search,
                            MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?) null,
                            MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?) null,
                            Dependencies,
                            Installed, Active, Required, Blocked, Latest, Location,
                            ForceScan,
                            messages).ContinueWith((FP) =>
                                                       {
                                                           if (FP.Result.Any())
                                                               foreach (Package P in FP.Result)
                                                                   output.Add(P);
                                                           else
                                                               Host.UI.WriteLine(ConsoleColor.Red,
                                                                                 Host.UI.RawUI.BackgroundColor,
                                                                                 "Empty list returned from GetPackages().");
                                                       }, TaskContinuationOptions.AttachedToParent);
                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet(VerbsCommon.Get, "PackageInfo", DefaultParameterSetName = "Typed")]
    public class GetPackageInfo : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Installed;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        protected override void ProcessRecord()
        {
                        // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Get-PackageInfo");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Installed != null) com.Parameters.Add("Installed", Installed);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                                    {
                                        foreach (Runspace R in Remotes)
                                        {
                                            Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                                                Remote.Exec(R, com));
                                            T.ContinueWith(antecedent =>
                                                    output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                                            T.Start();
                                        }
                                    });
                task.Start();
            }
            //Non-Remoting
            else
            {

                Package Details = null;
                PackageManagerMessages DetailMessages = new PackageManagerMessages()
                                                            {
                                                                PackageDetails = (P) =>
                                                                                     {
                                                                                         Details = P;
                                                                                     }
                                                            }.Extend(messages);
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.GetPackageDetails(InputPackage.CanonicalName, DetailMessages).ContinueWith((FP) =>
                                                                                                                 {
                                                                                                                     if
                                                                                                                         (
                                                                                                                         Details !=
                                                                                                                         null)
                                                                                                                         output
                                                                                                                             .
                                                                                                                             Add
                                                                                                                             (Details);
                                                                                                                 },
                                                                                                             TaskContinuationOptions
                                                                                                                 .
                                                                                                                 AttachedToParent);
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") +
                                        (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search,
                                              MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?) null,
                                              MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?) null,
                                              Dependencies, Installed, Active, Required, Blocked, Latest, Location,
                                              ForceScan, messages)
                                .ContinueWith((FP) =>
                                                    {
                                                        String P;
                                                        IEnumerable<Package> PL = FP.Result;
                                                        int i = 0;
                                                        if (PL.Count() > 1)
                                                        {
                                                            Collection<ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                                            foreach (Package p in PL)
                                                            {
                                                                string desc = "";
                                                                desc += Selections[i++];
                                                                desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                                Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                                            }
                                                            int choice =
                                                                Host.UI.PromptForChoice(
                                                                    "Multiple possible matches.",
                                                                    "Please choose one of the following:",
                                                                    Choices, 0);
                                                            P = Choices[choice].HelpMessage;
                                                            // do menu stuff here
                                                        }
                                                        else
                                                        {
                                                            P = PL.FirstOrDefault().CanonicalName;
                                                        }
                                                        PM.GetPackageDetails(P, DetailMessages)
                                                            .ContinueWith((FDP) =>
                                                                            {
                                                                                if (Details != null)
                                                                                    output.Add(Details);
                                                                            },
                                                                        TaskContinuationOptions.AttachedToParent);
                                                    });

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }
    }

    [Cmdlet(VerbsLifecycle.Install, "Package", DefaultParameterSetName = "Typed")]
    public class InstallPackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        [Parameter(Mandatory = false)]
        public bool? AutoUpgrade;
        [Parameter(Mandatory = false)]
        public SwitchParameter ForceInstall;
        [Parameter(Mandatory = false)]
        public SwitchParameter ForceDownload;
        [Parameter(Mandatory = false)]
        public SwitchParameter Pretend;
        

        // As this is a command to install something, I am disallowing the "Installed" switch.
        private bool? Installed = null;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Install-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                if (AutoUpgrade != null) com.Parameters.Add("AutoUpgrade", AutoUpgrade);
                if (ForceInstall) com.Parameters.Add("ForceInstall");
                if (ForceDownload) com.Parameters.Add("ForceDownload");
                if (Pretend) com.Parameters.Add("Pretend");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                int ID_Counter = 1;
                bool Running = true;
                AutoResetEvent Inbound = new AutoResetEvent(false);
                Dictionary<int, ProgressRecord> Recs = new Dictionary<int, ProgressRecord>();
                Dictionary<String, int> Children = new Dictionary<string, int>();
                PackageManagerMessages InstallMessages = new PackageManagerMessages()
                {
                    InstallingPackageProgress = (N, P, T) =>
                    {
                        if (!Children.ContainsKey(N))
                            Children[N] = ID_Counter++;
                        ProgressRecord current = new ProgressRecord(Children[N], N, "Installing...");
                        current.PercentComplete = P;
                        current.CurrentOperation = "Installing package...";
                        current.SecondsRemaining = -1;
                        Recs[Children[N]] = current;
                        Inbound.Set();
                    },
                    InstalledPackage = (N) =>
                    {
                        ProgressRecord current = new ProgressRecord(Children[N], N, "Installing...");
                        current.PercentComplete = 100;
                        current.CurrentOperation = "Installing package...";
                        current.SecondsRemaining = -1;
                        Recs[Children[N]] = current;
                        Inbound.Set(); 
                        Children.Remove(N);
                    },
                    FailedPackageInstall = (name,file,reason) => Host.UI.WriteLine("Installation failed for package: "+name+", from file "+file+", for reason: "+reason)
                }.Extend(messages);

                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.InstallPackage(CanonicalName, AutoUpgrade, ForceInstall, ForceDownload, Pretend, InstallMessages).ContinueWith(antecedent =>
                                                                                                                                                     {
                                                                                                                                                         Running = false;
                                                                                                                                                         Inbound.Set();
                                                                                                                                                     }, TaskContinuationOptions.AttachedToParent);
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.InstallPackage(P, AutoUpgrade, ForceInstall, ForceDownload, Pretend, InstallMessages);
                                       }, TaskContinuationOptions.AttachedToParent).ContinueWith(antecedent =>
                                                                                                     { 
                                                                                                         Running = false;
                                                                                                         Inbound.Set();
                                                                                                     });

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }

                do
                {
                    Inbound.WaitOne();
                    if (Recs.Any())
                        foreach (KeyValuePair<int, ProgressRecord> PR in Recs)
                        {
                            WriteProgress(PR.Value);
                            if (PR.Value.PercentComplete >= 100)
                                Recs.Remove(PR.Key);
                        }
                } while (Running);
        }
        }

    }

    [Cmdlet(VerbsCommon.Remove, "Package", DefaultParameterSetName = "Typed")]
    public class RemovePackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        [Parameter(Mandatory = false)]
        public SwitchParameter ForceRemove;
        [Parameter(Mandatory = false)]
        public SwitchParameter Pretend;


        // As this is a command to install something, I am disallowing the "Installed" switch.
        private bool? Installed = true;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Remove-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                if (ForceRemove) com.Parameters.Add("ForceRemove");
                if (Pretend) com.Parameters.Add("Pretend");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                PackageManagerMessages RemovalMessages = new PackageManagerMessages()
                {
                    FailedPackageRemoval = (N, R) => Host.UI.WriteLine("Failed to remove package:  " + N + ",  Reason: " + R)
                }.Extend(messages);
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.RemovePackage(CanonicalName, ForceRemove, RemovalMessages).ContinueWith((x) => Host.UI.WriteLine("Package removed: " + CanonicalName));
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.RemovePackage(P, ForceRemove, RemovalMessages).ContinueWith((x) => Host.UI.WriteLine("Package removed: " + P));
                                       }, TaskContinuationOptions.AttachedToParent);

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet(VerbsData.Update, "Package", DefaultParameterSetName = "Typed")]
    public class UpdatePackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        [Parameter(Mandatory = false)]
        public SwitchParameter Pretend;


        // As this is a command to install something, I am disallowing the "Installed" switch.
        private bool? Installed = true;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Update-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                if (Pretend) com.Parameters.Add("Pretend");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                Host.UI.WriteLine("Update-Package not yet implimented.");
            }

        }

    }

    [Cmdlet("Trim", "Packages")]
    public class TrimPackages : CoApp_Cmdlet
    {

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Trim-Packages");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                task = PM.GetPackages("*", null, null, null, true, null, false, false, null, null, null, messages).ContinueWith((Packs) =>
                                      {
                                          foreach (Package P in Packs.Result)
                                          {
                                              PM.RemovePackage(P.CanonicalName, null, messages).ContinueWith((A) => { }, TaskContinuationOptions.AttachedToParent);
                                          }
                                      }, TaskContinuationOptions.AttachedToParent);
            }
        }
    }

    [Cmdlet("Activate", "Package", DefaultParameterSetName = "Typed")]
    public class ActivatePackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        // We can only set a package as "Active" if it's already installed.
        private readonly bool? Installed = true;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Activate-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.SetPackage(CanonicalName, true, null, null, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Active': " + CanonicalName));
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.SetPackage(P, true, null, null, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Active': " + P));
                                       }, TaskContinuationOptions.AttachedToParent);

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet(VerbsSecurity.Block, "Package", DefaultParameterSetName = "Typed")]
    public class BlockPackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        private bool? Installed;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Block-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Installed != null) com.Parameters.Add("Installed", Installed);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.SetPackage(CanonicalName, null, null, true, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Blocked': " + CanonicalName));
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.SetPackage(P, null, null, true, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Blocked': " + P));
                                       }, TaskContinuationOptions.AttachedToParent);

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet(VerbsSecurity.Unblock, "Package", DefaultParameterSetName = "Typed")]
    public class UnblockPackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        private bool? Installed;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Unblock-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Installed != null) com.Parameters.Add("Installed", Installed);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.SetPackage(CanonicalName, null, null, false, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Unblocked': " + CanonicalName));
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.SetPackage(P, null, null, false, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Unblocked': " + P));
                                       }, TaskContinuationOptions.AttachedToParent);

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet("Mark", "Package", DefaultParameterSetName = "Typed")]
    public class MarkPackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        private bool? Installed;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Mark-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Installed != null) com.Parameters.Add("Installed", Installed);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.SetPackage(CanonicalName, null, true, null, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Required': " + CanonicalName));
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.SetPackage(P, null, true, null, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Required': " + P));
                                       }, TaskContinuationOptions.AttachedToParent);

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet("Unmark", "Package", DefaultParameterSetName = "Typed")]
    public class UnmarkPackage : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Package")]
        public Package InputPackage;

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Canonical")]
        public string CanonicalName;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Typed")]
        public string Name;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Version;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Arch;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string PublicKeyToken;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MinVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string MaxVersion;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Dependencies;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        private bool? Installed;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Active;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Required;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Blocked;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public bool? Latest;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? Index;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public int? MaxResults;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public string Location;
        [Parameter(Mandatory = false, ParameterSetName = "Typed")]
        public SwitchParameter ForceScan;

        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Unmark-Package");
                if (InputPackage != null) com.Parameters.Add("InputPackage", InputPackage);
                if (CanonicalName != null) com.Parameters.Add("CanonicalName", CanonicalName);
                if (Name != null) com.Parameters.Add("Name", Name);
                if (Version != null) com.Parameters.Add("Version", Version);
                if (Arch != null) com.Parameters.Add("Arch", Arch);
                if (PublicKeyToken != null) com.Parameters.Add("PublicKeyToken", PublicKeyToken);
                if (MinVersion != null) com.Parameters.Add("MinVersion", MinVersion);
                if (MaxVersion != null) com.Parameters.Add("MaxVersion", MaxVersion);
                if (Dependencies != null) com.Parameters.Add("Dependencies", Dependencies);
                if (Installed != null) com.Parameters.Add("Installed", Installed);
                if (Active != null) com.Parameters.Add("Active", Active);
                if (Required != null) com.Parameters.Add("Required", Required);
                if (Blocked != null) com.Parameters.Add("Blocked", Blocked);
                if (Latest != null) com.Parameters.Add("Latest", Latest);
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                if (Location != null) com.Parameters.Add("Location", Location);
                if (ForceScan) com.Parameters.Add("ForceScan");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                switch (ParameterSetName)
                {
                    case "Package":
                        CanonicalName = InputPackage.CanonicalName;
                        goto case "Canonical";
                    case "Canonical":
                        task = PM.SetPackage(CanonicalName, null, false, null, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Not Required': " + CanonicalName));
                        break;

                    case "Typed":
                        string search = Name + (Version != null ? "-" + Version : "") + (Arch != null ? "-" + Arch : "") + (PublicKeyToken != null ? "-" + PublicKeyToken : "");
                        task = PM.GetPackages(search, MinVersion != null ? MinVersion.VersionStringToUInt64() : (ulong?)null,
                                       MaxVersion != null ? MaxVersion.VersionStringToUInt64() : (ulong?)null,
                                       Dependencies, Installed, Active, Required, Blocked, Latest, Location, ForceScan,
                                       messages).ContinueWith((FP) =>
                                       {
                                           String P;
                                           IEnumerable<Package> PL = FP.Result;
                                           int i = 0;
                                           if (PL.Count() > 1)
                                           {
                                               Collection<System.Management.Automation.Host.ChoiceDescription> Choices = new Collection<ChoiceDescription>();
                                               foreach (Package p in PL)
                                               {
                                                   string desc = "";
                                                   desc += Selections[i++];
                                                   desc += p.Name + "-" + p.Version + "-" + p.Architecture;
                                                   Choices.Add(new ChoiceDescription(desc, p.CanonicalName));
                                               }
                                               int choice = Host.UI.PromptForChoice("Multiple possible matches.",
                                                                       "Please choose one of the following:", Choices, 0);
                                               P = Choices[choice].HelpMessage;
                                               // do menu stuff here
                                           }
                                           else
                                           {
                                               P = PL.FirstOrDefault().CanonicalName;
                                           }
                                           PM.SetPackage(P, null, false, null, messages).ContinueWith((x) => Host.UI.WriteLine("Package set to 'Not Required': " + P));
                                       }, TaskContinuationOptions.AttachedToParent);

                        break;
                    default:
                        Host.UI.WriteLine("Invalid input parameters.");
                        break;

                }
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "Feed")]
    public class AddFeed : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string Location;
        [Parameter(Mandatory = false)]
        public SwitchParameter SessionOnly;


        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Add-Feed");
                if (Location != null) com.Parameters.Add("Location", Location);
                if (SessionOnly) com.Parameters.Add("SessionOnly");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() => Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                PackageManagerMessages FeedMessages = new PackageManagerMessages()
                {
                    FeedAdded = (S) => Host.UI.WriteLine("Feed Added:  " + S)
                }.Extend(messages);
                task = PM.AddFeed(Location, SessionOnly, FeedMessages);
            }
        }
    }

    [Cmdlet(VerbsCommon.Remove, "Feed")]
    public class RemoveFeed : CoApp_Cmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string Location;
        [Parameter(Mandatory = false)]
        public SwitchParameter SessionOnly;


        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("Remove-Feed");
                if (Location != null) com.Parameters.Add("Location", Location);
                if (SessionOnly) com.Parameters.Add("SessionOnly");
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                PackageManagerMessages FeedMessages = new PackageManagerMessages()
                {
                    FeedRemoved = (S) => Host.UI.WriteLine("Feed Removed:  " + S)
                }.Extend(messages);
                task = PM.RemoveFeed(Location, SessionOnly, FeedMessages);
            }
        }
    }

    [Cmdlet("List", "Feed")]
    public class ListFeed : CoApp_Cmdlet
    {
        [Parameter(Mandatory = false)]
        public int? Index;
        [Parameter(Mandatory = false)]
        public int? MaxResults;


        protected override void ProcessRecord()
        {
            // Handle remoting
            if (!ComputerName.IsNullOrEmpty())
            {
                Command com = new Command("List-Feed");
                if (Index != null) com.Parameters.Add("Index", Index);
                if (MaxResults != null) com.Parameters.Add("MaxResults", MaxResults);
                task = new Task(() =>
                {
                    foreach (Runspace R in Remotes)
                    {
                        Task<Collection<PSObject>> T = new Task<Collection<PSObject>>(() =>
                            Remote.Exec(R, com));
                        T.ContinueWith(antecedent =>
                                output.AddRange(antecedent.Result), TaskContinuationOptions.AttachedToParent);
                        T.Start();
                    }
                });
                task.Start();
            }
            //Non-Remoting
            else
            {
                PackageManagerMessages FeedMessages = new PackageManagerMessages()
                {

                    FeedDetails = (Loc, Scanned, isSession, supressed, verified) => output
                        .Add(new Feed()
                                {
                                    Location = Loc,
                                    LastScanned = Scanned,
                                    Session = isSession,
                                    Supressed = supressed,
                                    Verified = verified
                                })
                }.Extend(messages);
                task = PM.ListFeeds(Index, MaxResults, FeedMessages);

            }
        }
    }



    /*
        [Cmdlet(VerbsData.ConvertTo, "String")]
        public class ConvertTo_String : PSCmdlet
        {
            [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
            public byte[] Bytes;



            protected override void ProcessRecord()
            {
                string S = CommonMethods.Bytes_To_String(Bytes);
                WriteObject(S);
            }

        }

        */
}

