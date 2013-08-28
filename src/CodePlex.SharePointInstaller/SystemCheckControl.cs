/**********************************************************************/
/*                                                                    */
/*                   SharePoint Solution Installer                    */
/*             http://www.codeplex.com/sharepointinstaller            */
/*                                                                    */
/*               (c) Copyright 2007 Lars Fastrup Nielsen.             */
/*                                                                    */
/*  This source is subject to the Microsoft Permissive License.       */
/*  http://www.codeplex.com/sharepointinstaller/Project/License.aspx  */
/*                                                                    */
/* KML: Minor update to usage of EULA config property and error text  */
/*                                                                    */
/**********************************************************************/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Security;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Timers;
using System.Windows.Forms;

using Microsoft.Win32;

using Microsoft.SharePoint.Administration;
using System.Configuration;
using System.IO;
using CodePlex.SharePointInstaller.Resources;


namespace CodePlex.SharePointInstaller
{
    public enum SharePointVersion
    {
        SP2007,
        SP2010,
        Unrestricted
    }

    public partial class SystemCheckControl : InstallerControl
    {
        private static readonly ILog log = LogManager.GetLogger();
        private System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

        private bool requireMOSS;
        private bool requireSearchSKU;
        private SystemCheckList checks;
        private int nextCheckIndex;
        private int errors;

        #region Constructors

        public SystemCheckControl()
        {
            InitializeComponent();

            this.Load += new EventHandler(SystemCheckControl_Load);
        }

        #endregion

        #region Public Properties

        public SharePointVersion SupportedSharePointVersion
        {
            set;
            get;
        }

        public bool RequireMOSS
        {
            get { return requireMOSS; }
            set { requireMOSS = value; }
        }

        public bool RequireSearchSKU
        {
            get { return requireSearchSKU; }
            set { requireSearchSKU = value; }
        }

        #endregion

        #region Event Handlers

        private void SystemCheckControl_Load(object sender, EventArgs e)
        {
        }

        private void TimerEventProcessor(Object myObject, EventArgs myEventArgs)
        {
            timer.Stop();

            if (nextCheckIndex < checks.Count)
            {
                if (ExecuteCheck(nextCheckIndex))
                {
                    nextCheckIndex++;
                    timer.Start();
                    return;
                }
            }

            FinalizeChecks();
        }

        #endregion

        #region Protected Methods

        protected internal override void Open(InstallOptions options)
        {
            if (checks == null)
            {
                Form.NextButton.Enabled = false;
                Form.PrevButton.Enabled = false;

                checks = new SystemCheckList();
                InitializeChecks();

                timer.Interval = 100;
                timer.Tick += new EventHandler(TimerEventProcessor);
                timer.Start();
            }
        }

        protected internal override void Close(InstallOptions options)
        {
        }

        #endregion

        #region Private Methods

        private void InitializeChecks()
        {
            this.tableLayoutPanel.SuspendLayout();

            //
            // WSS Installed Check
            //
            AddCheck(GetWssOrSfInstalledCheck());

            //
            // MOSS Installed Check
            //
            if (requireMOSS)
            {
                AddCheck(GetMossOrSharePoint2010InstalledCheck());
            }

            //
            // Admin Rights Check
            //
            AdminRightsCheck adminRightsCheck = new AdminRightsCheck();
            adminRightsCheck.QuestionText = CommonUIStrings.adminRightsCheckQuestionText;
            adminRightsCheck.OkText = CommonUIStrings.adminRightsCheckOkText;
            adminRightsCheck.ErrorText = CommonUIStrings.adminRightsCheckErrorText;
            AddCheck(adminRightsCheck);

            //
            // Admin Service Check
            //
            //AdminServiceCheck adminServiceCheck = new AdminServiceCheck();
            //adminServiceCheck.QuestionText = CommonUIStrings.adminServiceCheckQuestionText;
            //adminServiceCheck.OkText = CommonUIStrings.adminServiceCheckOkText;
            //adminServiceCheck.ErrorText = CommonUIStrings.adminServiceCheckErrorText;
            //AddCheck(adminServiceCheck);

            //
            // Timer Service Check
            //
            AddCheck(GetTimerRunningCheck());

            //
            // Solution File Check
            //
            SolutionFileCheck solutionFileCheck = new SolutionFileCheck();
            solutionFileCheck.QuestionText = CommonUIStrings.solutionFileCheckQuestionText;
            solutionFileCheck.OkText = CommonUIStrings.solutionFileCheckOkText;
            AddCheck(solutionFileCheck);

            //
            // Solution Check
            //
            SolutionCheck solutionCheck = new SolutionCheck();
            solutionCheck.QuestionText = InstallConfiguration.FormatString(CommonUIStrings.solutionCheckQuestionText);
            solutionCheck.OkText = InstallConfiguration.FormatString(CommonUIStrings.solutionFileCheckOkText);
            solutionCheck.ErrorText = InstallConfiguration.FormatString(CommonUIStrings.solutionCheckErrorText);
            AddCheck(solutionCheck);

            //
            // Add empty row that will eat up the rest of the 
            // row space in the layout table.
            //
            this.tableLayoutPanel.RowCount++;
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));

            this.tableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel.PerformLayout();
        }

        private SystemCheck GetTimerRunningCheck()
        {
            switch (SupportedSharePointVersion)
            {
                case SharePointVersion.SP2007:
                    {
                        return GetTimerCheck(12);
                    }
                case SharePointVersion.SP2010:
                    {
                        return GetTimerCheck(14);
                    }
                default:
                    {
                        BinaryCheck binaryCheck = new BinaryCheck(GetTimerCheck(12), GetTimerCheck(14));
                        binaryCheck.QuestionText = CommonUIStrings.mossOrSharePoint2010QuestionText;
                        return binaryCheck;
                    }
            }
        }


        private SystemCheck GetTimerCheck(int version)
        {
            TimerServiceCheck timerServiceCheck = new TimerServiceCheck(version);
            timerServiceCheck.QuestionText = CommonUIStrings.timerServiceCheckQuestionText;
            timerServiceCheck.OkText = CommonUIStrings.timerServiceCheckOkText;
            timerServiceCheck.ErrorText = CommonUIStrings.timerServiceCheckErrorText;
            return timerServiceCheck;
        }

        private SystemCheck GetMossCheck()
        {
            MOSSInstalledCheck mossCheck = new MOSSInstalledCheck(12);
            mossCheck.QuestionText = CommonUIStrings.mossCheckQuestionText;
            mossCheck.OkText = CommonUIStrings.mossCheckOkText;
            mossCheck.ErrorText = CommonUIStrings.mossCheckErrorText;
            return mossCheck;
        }

        private SystemCheck GetSP2010Check()
        {
            MOSSInstalledCheck mossCheck = new MOSSInstalledCheck(14);
            mossCheck.QuestionText = CommonUIStrings.sharePoint2010QuestionText;
            mossCheck.OkText = CommonUIStrings.sharePoint2010OkText;
            mossCheck.ErrorText = CommonUIStrings.sharePoint2010ErrorText;
            return mossCheck;
        }

        private SystemCheck GetMossOrSharePoint2010InstalledCheck()
        {
            switch (SupportedSharePointVersion)
            {
                case SharePointVersion.SP2007:
                    {
                        return GetMossCheck();
                    }
                case SharePointVersion.SP2010:
                    {
                        return GetSP2010Check();
                    }
                default:
                    {
                        BinaryCheck binaryCheck = new BinaryCheck(GetMossCheck(), GetSP2010Check());
                        binaryCheck.QuestionText = CommonUIStrings.mossOrSharePoint2010QuestionText;
                        return binaryCheck;
                    }
            }
        }

        private SystemCheck GetWssOrSfInstalledCheck()
        {
            switch (SupportedSharePointVersion)
            {
                case SharePointVersion.SP2007:
                    {
                        return GetWssInstalledCheck();
                    }
                case SharePointVersion.SP2010:
                    {
                        return GetSfInstalledCheck();
                    }
                //all versions are supported by default
                default:
                    {
                        BinaryCheck binaryCheck = new BinaryCheck(GetWssInstalledCheck(), GetSfInstalledCheck());
                        binaryCheck.QuestionText = CommonUIStrings.wssOrSfCheckQuestionText;
                        return binaryCheck;
                    }
            }
        }

        private static SystemCheck GetSfInstalledCheck()
        {
            WSSInstalledCheck sfCheck = new WSSInstalledCheck(14);
            sfCheck.QuestionText = CommonUIStrings.sfCheckQuestionText;
            sfCheck.OkText = CommonUIStrings.sfCheckOkText;
            sfCheck.ErrorText = CommonUIStrings.sfCheckErrorText;
            return sfCheck;
        }

        private static SystemCheck GetWssInstalledCheck()
        {
            WSSInstalledCheck wssCheck = new WSSInstalledCheck(12);
            wssCheck.QuestionText = CommonUIStrings.wssCheckQuestionText;
            wssCheck.OkText = CommonUIStrings.wssCheckOkText;
            wssCheck.ErrorText = CommonUIStrings.wssCheckErrorText;
            return wssCheck;
        }

        private bool ExecuteCheck(int index)
        {
            SystemCheck check = checks[index];
            string imageLabelName = "imageLabel" + index;
            string textLabelName = "textLabel" + index;
            Label imageLabel = (Label)tableLayoutPanel.Controls[imageLabelName];
            Label textLabel = (Label)tableLayoutPanel.Controls[textLabelName];

            try
            {
                SystemCheckResult result = check.Execute();
                if (result == SystemCheckResult.Success)
                {
                    imageLabel.Image = global::CodePlex.SharePointInstaller.Properties.Resources.CheckOk;
                    textLabel.Text = check.OkText;
                }
                else if (result == SystemCheckResult.Error)
                {
                    errors++;
                    imageLabel.Image = global::CodePlex.SharePointInstaller.Properties.Resources.CheckFail;
                    textLabel.Text = check.ErrorText;
                }

                //
                // Show play icon on next check that will run.
                //
                int nextIndex = index + 1;
                string nextImageLabelName = "imageLabel" + nextIndex;
                Label nextImageLabel = (Label)tableLayoutPanel.Controls[nextImageLabelName];
                if (nextImageLabel != null)
                {
                    nextImageLabel.Image = global::CodePlex.SharePointInstaller.Properties.Resources.CheckPlay;
                }

                return true;
            }

            catch (InstallException ex)
            {
                errors++;
                imageLabel.Image = global::CodePlex.SharePointInstaller.Properties.Resources.CheckFail;
                textLabel.Text = ex.Message;
            }

            return false;
        }

        private void FinalizeChecks()
        {
            if (errors == 0)
            {
                ConfigureControls();
                Form.NextButton.Enabled = true;
                messageLabel.Text = CommonUIStrings.messageLabelTextSuccess;
            }
            else
            {
                messageLabel.Text = InstallConfiguration.FormatString(CommonUIStrings.messageLabelTextError);
            }

            Form.PrevButton.Enabled = true;
        }

        private void AddCheck(SystemCheck check)
        {
            int row = tableLayoutPanel.RowCount;

            Label imageLabel = new Label();
            imageLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            imageLabel.Image = global::CodePlex.SharePointInstaller.Properties.Resources.CheckWait;
            imageLabel.Location = new System.Drawing.Point(3, 0);
            imageLabel.Name = "imageLabel" + row;
            imageLabel.Size = new System.Drawing.Size(24, 20);

            Label textLabel = new Label();
            textLabel.AutoSize = true;
            textLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            textLabel.Location = new System.Drawing.Point(33, 0);
            textLabel.Name = "textLabel" + row;
            textLabel.Size = new System.Drawing.Size(390, 20);
            textLabel.Text = check.QuestionText;
            textLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            this.tableLayoutPanel.Controls.Add(imageLabel, 0, row);
            this.tableLayoutPanel.Controls.Add(textLabel, 1, row);
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel.RowCount++;

            checks.Add(check);
        }

        private void ConfigureControls()
        {
            SolutionCheck check = (SolutionCheck)checks["SolutionCheck"];
            SPSolution solution = check.Solution;

            if (solution == null)
            {
                AddInstallControls();
            }

            else
            {
                Version installedVersion = InstallConfiguration.InstalledVersion;
                Version newVersion = InstallConfiguration.SolutionVersion;

                if (newVersion != installedVersion)
                {
                    Form.ContentControls.Add(Program.CreateUpgradeControl());
                }
                else
                {
                    Form.ContentControls.Add(Program.CreateRepairControl());
                }
            }
        }

        private void AddInstallControls()
        {
            //
            // Add EULA control if an EULA file was specified.
            //
            string filename = InstallConfiguration.EULA;
            if (!String.IsNullOrEmpty(filename))
            {
                Form.ContentControls.Add(Program.CreateEULAControl());
            }

            Form.ContentControls.Add(Program.CreateDeploymentTargetsControl());
            //Form.ContentControls.Add(Program.CreateOptionsControl());
            Form.ContentControls.Add(Program.CreateProcessControl());
        }

        #endregion

        #region Check Classes

        private enum SystemCheckResult
        {
            Inconclusive,
            Success,
            Error
        }

        /// <summary>
        /// Base class for all system checks.
        /// </summary>
        private abstract class SystemCheck
        {
            private readonly string id;
            private string questionText;
            private string okText;
            private string errorText;

            protected SystemCheck(string id)
            {
                this.id = id;
            }

            public string Id
            {
                get { return id; }
            }

            public string QuestionText
            {
                get { return questionText; }
                set { questionText = value; }
            }

            public string OkText
            {
                get { return okText; }
                set { okText = value; }
            }

            public string ErrorText
            {
                get { return errorText; }
                set { errorText = value; }
            }

            internal SystemCheckResult Execute()
            {
                if (CanRun)
                {
                    return DoExecute();
                }
                return SystemCheckResult.Inconclusive;
            }

            protected abstract SystemCheckResult DoExecute();

            protected virtual bool CanRun
            {
                get { return true; }
            }

            protected static bool IsWSSInstalled()
            {
                return IsWSSInstalled(12) || IsWSSInstalled(14);
            }

            protected static bool IsWSSInstalled(int version)
            {
                string path = string.Format(@"SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\{0}.0", version);
                try
                {
                    RegistryKey key = Registry.LocalMachine.OpenSubKey(path);
                    if (key != null)
                    {
                        object val = key.GetValue("SharePoint");
                        if (val != null && val.Equals("Installed"))
                        {
                            return true;
                        }
                    }
                    return false;
                }

                catch (SecurityException ex)
                {
                    throw new InstallException(string.Format(Resources.CommonUIStrings.installExceptionAccessDenied, path), ex);
                }
            }

            protected static bool IsMOSSInstalled(int version)
            {
                string name = string.Format(@"SOFTWARE\Microsoft\Office Server\{0}.0", version);

                try
                {
                    RegistryKey key = Registry.LocalMachine.OpenSubKey(name);
                    if (key != null)
                    {
                        string versionStr = key.GetValue("BuildVersion") as string;
                        if (versionStr != null)
                        {
                            Version buildVersion = new Version(versionStr);
                            if (buildVersion.Major == version)
                            {
                                return true;
                            }
                        }
                    }
                    return false;
                }

                catch (SecurityException ex)
                {
                    throw new InstallException(string.Format(CommonUIStrings.installExceptionAccessDenied, name), ex);
                }
            }
        }

        private class SystemCheckList : List<SystemCheck>
        {
            internal SystemCheck this[string id]
            {
                get
                {
                    foreach (SystemCheck check in this)
                    {
                        if (check.Id == id) return check;
                    }
                    return null;
                }
            }
        }

        /// <summary>
        /// Checks if WSS 3.0 is installed.
        /// </summary>
        private class WSSInstalledCheck : SystemCheck
        {
            int version;
            internal WSSInstalledCheck(int version) : base("WSSInstalledCheck") { this.version = version; }

            protected override SystemCheckResult DoExecute()
            {
                if (IsWSSInstalled(version)) return SystemCheckResult.Success;
                return SystemCheckResult.Error;
            }
        }

        /// <summary>
        /// Checks if one of checks is successful is installed.
        /// </summary>
        private class BinaryCheck : SystemCheck
        {
            SystemCheck first;
            SystemCheck second;
            internal BinaryCheck(SystemCheck first, SystemCheck second)
                : base("BinaryCheck")
            {
                if (first == null)
                    throw new ArgumentNullException("first");
                if (second == null)
                    throw new ArgumentNullException("second");
                this.first = first;
                this.second = second;
            }

            protected override SystemCheckResult DoExecute()
            {
                SystemCheckResult firstResult = first.Execute();
                SystemCheckResult secondResult = second.Execute();
                if (firstResult == SystemCheckResult.Success)
                {
                    OkText = first.OkText;
                    return SystemCheckResult.Success;
                }
                else
                {
                    if (secondResult == SystemCheckResult.Success)
                    {
                        OkText = second.OkText;
                        return SystemCheckResult.Success;
                    }
                    else
                    {
                        ErrorText = first.ErrorText;
                        return firstResult;
                    }
                }
            }
        }

        /// <summary>
        /// Checks if Microsoft Office Server 2007 is installed.
        /// </summary>
        private class MOSSInstalledCheck : SystemCheck
        {
            int version;
            internal MOSSInstalledCheck(int version)
                : base("MOSSInstalledCheck")
            {
                this.version = version;
            }

            protected override SystemCheckResult DoExecute()
            {
                if (IsMOSSInstalled(version)) return SystemCheckResult.Success;
                return SystemCheckResult.Error;
            }
        }

        /// <summary>
        /// Checks if the Windows SharePoint Services Administration service is started.
        /// </summary>
        private class AdminServiceCheck : SystemCheck
        {
            internal AdminServiceCheck() : base("AdminServiceCheck") { }

            protected override SystemCheckResult DoExecute()
            {
                try
                {
                    ServiceController sc = new ServiceController("SPAdmin");
                    if (sc.Status == ServiceControllerStatus.Running)
                    {
                        return SystemCheckResult.Success;
                    }
                    return SystemCheckResult.Error;
                }

                catch (Win32Exception ex)
                {
                    log.Error(ex.Message, ex);
                }

                catch (InvalidOperationException ex)
                {
                    log.Error(ex.Message, ex);
                }

                return SystemCheckResult.Inconclusive;
            }

            protected override bool CanRun
            {
                get { return IsWSSInstalled(); }
            }
        }

        /// <summary>
        /// Checks if the Windows SharePoint Services Timer service is started.
        /// </summary>
        private class TimerServiceCheck : SystemCheck
        {
            int version;
            internal TimerServiceCheck(int version)
                : base("TimerServiceCheck")
            {
                this.version = version;
            }

            protected string GetTimerNameForVersion(int version)
            {
                switch (version)
                {
                    case 12:
                        return "SPTimerV3";
                    case 14:
                        return "SPTimerV4";
                    default:
                        throw new Exception(string.Format("No timer name defined for version {0}", version));

                }
            }
            protected override SystemCheckResult DoExecute()
            {
                try
                {
                    string timerName = GetTimerNameForVersion(version);
                    ServiceController sc = new ServiceController(timerName);
                    if (sc.Status == ServiceControllerStatus.Running)
                    {
                        return SystemCheckResult.Success;
                    }
                    return SystemCheckResult.Error;

                    //
                    // LFN 2009-06-21: Do not restart the time service anymore. First it does
                    // not always work with Windows Server 2008 where it seems a local 
                    // admin may not necessarily be allowed to start and stop the service.
                    // Secondly, the timer service has become more stable with WSS SP1 and SP2.
                    //
                    /*TimeSpan timeout = new TimeSpan(0, 0, 60);
                    ServiceController sc = new ServiceController("SPTimerV3");
                    if (sc.Status == ServiceControllerStatus.Running)
                    {
                      sc.Stop();
                      sc.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
                    }

                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, timeout);

                    return SystemCheckResult.Success;*/
                }

                catch (System.ServiceProcess.TimeoutException ex)
                {
                    log.Error(ex.Message, ex);
                }

                catch (Win32Exception ex)
                {
                    log.Error(ex.Message, ex);
                }

                catch (InvalidOperationException ex)
                {
                    log.Error(ex.Message, ex);
                    return SystemCheckResult.Error;
                }

                return SystemCheckResult.Inconclusive;
            }

            protected override bool CanRun
            {
                get { return IsWSSInstalled(); }
            }
        }

        /// <summary>
        /// Checks if the current user is an administrator.
        /// </summary>
        private class AdminRightsCheck : SystemCheck
        {
            internal AdminRightsCheck() : base("AdminRightsCheck") { }

            protected override SystemCheckResult DoExecute()
            {
                try
                {
                    if (SPFarm.Local.CurrentUserIsAdministrator())
                    {
                        return SystemCheckResult.Success;
                    }
                    return SystemCheckResult.Error;
                }

                catch (NullReferenceException)
                {
                    throw new InstallException(CommonUIStrings.installExceptionDatabase);
                }

                catch (Exception ex)
                {
                    throw new InstallException(ex.Message, ex);
                }
            }

            protected override bool CanRun
            {
                get { return IsWSSInstalled(); }
            }
        }

        private class SolutionFileCheck : SystemCheck
        {
            internal SolutionFileCheck() : base("SolutionFileCheck") { }

            protected override SystemCheckResult DoExecute()
            {
                string filename = InstallConfiguration.SolutionFile;
                if (!String.IsNullOrEmpty(filename))
                {
                    FileInfo solutionFileInfo = new FileInfo(filename);
                    if (!solutionFileInfo.Exists)
                    {
                        throw new InstallException(string.Format(CommonUIStrings.installExceptionFileNotFound, filename));
                    }
                }
                else
                {
                    throw new InstallException(CommonUIStrings.installExceptionConfigurationNoWsp);
                }

                return SystemCheckResult.Success;
            }
        }

        private class SolutionCheck : SystemCheck
        {
            private SPSolution solution;

            internal SolutionCheck() : base("SolutionCheck") { }

            protected override SystemCheckResult DoExecute()
            {
                Guid solutionId = Guid.Empty;
                try
                {
                    solutionId = InstallConfiguration.SolutionId;
                }
                catch (ArgumentNullException)
                {
                    throw new InstallException(CommonUIStrings.installExceptionConfigurationNoId);
                }
                catch (FormatException)
                {
                    throw new InstallException(CommonUIStrings.installExceptionConfigurationInvalidId);
                }

                try
                {
                    solution = SPFarm.Local.Solutions[solutionId];
                    if (solution != null)
                    {
                        this.OkText = InstallConfiguration.FormatString(CommonUIStrings.solutionCheckOkTextInstalled);
                    }
                    else
                    {
                        this.OkText = InstallConfiguration.FormatString(CommonUIStrings.solutionCheckOkTextNotInstalled);
                    }
                }

                catch (NullReferenceException)
                {
                    throw new InstallException(CommonUIStrings.installExceptionDatabase);
                }

                catch (Exception ex)
                {
                    throw new InstallException(ex.Message, ex);
                }

                return SystemCheckResult.Success;
            }

            protected override bool CanRun
            {
                get { return IsWSSInstalled(); }
            }

            internal SPSolution Solution
            {
                get { return solution; }
            }
        }

        #endregion
    }
}
