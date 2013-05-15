using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using GoContactSyncMod.Properties;
using System.Threading;

namespace GoContactSyncMod
{
    internal partial class SettingsForm : Form
    {
        private static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);
        private static readonly int WM_GCSM_SHOWME = RegisterWindowMessage("WM_GCSM_SHOWME");
        private static readonly TimeSpan BalloonTimeout = TimeSpan.FromSeconds(5);
        private static readonly Icon[] WorkIcons = new Icon[] { 
            Resources.Work_01, 
            Resources.Work_02, 
            Resources.Work_03,
            Resources.Work_04,
            Resources.Work_05,
            Resources.Work_06,
            Resources.Work_07, 
            Resources.Work_08,
            Resources.Work_09, 
            Resources.Work_10, 
            Resources.Work_11, 
            Resources.Work_12 
        };

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int RegisterWindowMessage(string message);

        [Flags]
        private enum WorkerTasks
        {
            None = 0,
            ResetMatches = 1,
            Synchronize = 2,
        }

        private class SyncContext
        {
            private DateTime lastImportantReport;
            private string statusText;
            private ToolTipIcon statusIcon;

            public SyncContext(Settings settings, WorkerTasks task, bool interactive)
            {
                // check the arguments and store the values
                if (settings == null)
                    throw new ArgumentNullException("settings");
                if (settings.IsFirstSync)
                    task |= WorkerTasks.ResetMatches;
                Tasks = task;
                UserName = settings.UserName;
                Password = settings.Password;
                Mode = settings.SyncMode;
                Interactive = interactive;
            }

            public WorkerTasks Tasks { get; private set; }
            public string UserName { get; private set; }
            public byte[] Password { get; private set; }
            public SyncOption Mode { get; private set; }
            public bool Interactive { get; private set; }

            public void GetLastReport(out string text, out ToolTipIcon icon)
            {
                // make sure that there has been a report and return it
                if (statusText == null)
                    throw new InvalidOperationException();
                lock (this)
                {
                    text = statusText;
                    icon = statusIcon;
                }
            }

            public void Report(BackgroundWorker worker, string text, ToolTipIcon icon, bool isImportant = false)
            {
                // check the input values
                if (worker == null)
                    throw new ArgumentNullException("worker");
                if (text == null)
                    throw new ArgumentNullException("text");

                // wait for important reports to go away if we operate interactively
                if (Interactive)
                {
                    var timespanSinceLastImportantReport = DateTime.Now - lastImportantReport;
                    if (timespanSinceLastImportantReport < BalloonTimeout)
                        Thread.Sleep(BalloonTimeout - timespanSinceLastImportantReport);
                    if (isImportant)
                        lastImportantReport = DateTime.Now;
                }

                // make these changes within a lock and report the progress
                lock (this)
                {
                    statusText = text;
                    statusIcon = icon;
                }
                worker.ReportProgress(0, this);
            }
        }

        public static void ShowRemote()
        {
            // send a broadcast message to show the settings form in another instance
            PostMessage(HWND_BROADCAST, WM_GCSM_SHOWME, IntPtr.Zero, IntPtr.Zero);
        }

        private static byte[] EncodePassword(string plain)
        {
            // encrypt the plain password
            return string.IsNullOrEmpty(plain) ? null : ProtectedData.Protect(Encoding.Unicode.GetBytes(plain), null, DataProtectionScope.CurrentUser);
        }

        private static string DecodePassword(byte[] encrypted)
        {
            // decrypt the given password
            try { return (encrypted == null || encrypted.Length == 0) ? string.Empty : Encoding.Unicode.GetString(ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser)); }
            catch (CryptographicException) { return string.Empty; }
        }

        public SettingsForm()
        {
            // create all components, listeners and bindings
            InitializeComponent();
            Settings.Default.PropertyChanged += new PropertyChangedEventHandler(Settings_PropertyChanged);
            CreateBindings();

            // set the proper title texts
            Text = string.Format(Text, Application.ProductVersion);
            Notifications.BalloonTipTitle = Text;

            // do the stuff that is usually done by event handlers
            UpdateNotificationStatus(string.Empty, ToolTipIcon.None, false);
            UpdateWorkerStatus();
            UpdateSaveStatus();

            // try to upgrade the settings if there's no user name
            if (string.IsNullOrEmpty(Settings.Default.UserName))
                Settings.Default.Upgrade();
        }

        private void CreateBindings()
        {
            // bind the user name unmodified
            UserName.DataBindings.Add(new Binding("Text", Settings.Default, "UserName", false, DataSourceUpdateMode.OnPropertyChanged, string.Empty));

            // bind the password to the encrypted password
            var passwordBinding = new Binding("Text", Settings.Default, "Password", true, DataSourceUpdateMode.OnPropertyChanged);
            passwordBinding.Parse += (sender, e) => e.Value = EncodePassword((string)e.Value);
            passwordBinding.Format += (sender, e) => e.Value = DecodePassword((byte[])e.Value);
            Password.DataBindings.Add(passwordBinding);

            // bind the visibility of the sign-up link to the lack of a user name
            var signupBinding = new Binding("Visible", Settings.Default, "UserName", true, DataSourceUpdateMode.Never);
            signupBinding.Format += (sender, e) => e.Value = string.IsNullOrEmpty((string)e.Value);
            GoogleContactsSignup.DataBindings.Add(signupBinding);

            // bind all radio buttons
            CreateOptionBinding(TwoWaySync, SyncOption.MergeOutlookWins);
            CreateOptionBinding(GoogleToOutlook, SyncOption.GoogleToOutlookOnly);
            CreateOptionBinding(OutlookToGoogle, SyncOption.OutlookToGoogleOnly);

            // bind the interval control
            var syncIntervalBinding = new Binding("Value", Settings.Default, "SyncInterval", true, DataSourceUpdateMode.OnPropertyChanged);
            syncIntervalBinding.Parse += (sender, e) => e.Value = TimeSpan.FromMinutes((int)((decimal)e.Value));
            syncIntervalBinding.Format += (sender, e) => e.Value = (decimal)(int)((TimeSpan)e.Value).TotalMinutes;
            SyncInterval.DataBindings.Add(syncIntervalBinding);
        }

        private void CreateOptionBinding(RadioButton button, SyncOption mode)
        {
            // store the mode in the tag
            button.Tag = mode;

            // one-way sync the state from the data source with the check state
            var modeBinding = new Binding("Checked", Settings.Default, "SyncMode", true, DataSourceUpdateMode.Never);
            modeBinding.Format += (sender, e) => e.Value = (SyncOption)((Binding)sender).Control.Tag == (SyncOption)e.Value;
            button.DataBindings.Add(modeBinding);

            // update the data source when the check box is clicked (and don't autocheck it)
            button.AutoCheck = false;
            button.Click += (sender, e) => Settings.Default.SyncMode = (SyncOption)((RadioButton)sender).Tag;
        }

        protected override void WndProc(ref Message m)
        {
            // display the current form if we're asked to
            if (m.Msg == WM_GCSM_SHOWME)
            {
                Show();
                Activate();
            }
            base.WndProc(ref m);
        }

        private void UpdateNotificationStatus(string balloonText, ToolTipIcon balloonIcon, bool showBalloon)
        {
            // generate the hint text and make it shorter if necessary
            var text = string.IsNullOrEmpty(balloonText) ? Text : (Text + Environment.NewLine + balloonText);
            while (text.Length >= 64)
            {
                var newLine = text.LastIndexOf(Environment.NewLine);
                if (newLine == -1)
                {
                    text = text.Substring(0, 60) + "...";
                    break;
                }
                text = text.Substring(0, newLine);
            }

            // assign the values
            Notifications.Text = text;
            Notifications.BalloonTipText = balloonText;
            Notifications.BalloonTipIcon = balloonIcon;

            // show the balloon if requested
            if (showBalloon)
                Notifications.ShowBalloonTip((int)BalloonTimeout.TotalMilliseconds);
        }

        private void Sync(bool onlyResetMatches, bool interactive)
        {
            // don't do nothin' if we're already syncing
            if (Worker.IsBusy)
            {
                if (interactive)
                    Notifications.ShowBalloonTip((int)BalloonTimeout.TotalMilliseconds, Text, Resources.SettingsForm_SyncPending, ToolTipIcon.Info);
                return;
            }

            // ensure the settings aren't dirty
            if (Settings.Default.IsDirty)
            {
                if (interactive)
                {
                    // activate the form and the most appropriate button and show a message to the user
                    Activate();
                    if (Save.Enabled)
                        Save.Focus();
                    else
                        Cancel.Focus();
                    Notifications.ShowBalloonTip((int)BalloonTimeout.TotalMilliseconds, Text, Resources.SettingsForm_UnsavedSettings, ToolTipIcon.Info);
                }
                return;
            }

            // only continue if a user name was provided
            if (string.IsNullOrEmpty(Settings.Default.UserName))
            {
                if (interactive)
                {
                    // show and activate the form, focus the user name input and show a message to the user
                    Show();
                    Activate();
                    UserName.Focus();
                    Notifications.ShowBalloonTip((int)BalloonTimeout.TotalMilliseconds, Text, Resources.SettingsForm_SettingsIncomplete, ToolTipIcon.Info);
                }
                return;
            }

            // start the worker and update the UI
            Worker.RunWorkerAsync(new SyncContext(Settings.Default, onlyResetMatches ? WorkerTasks.ResetMatches : WorkerTasks.Synchronize, interactive));
            UpdateWorkerStatus();
        }

        private void UpdateSaveStatus()
        {
            // set the enabled state of the save button
            Save.Enabled = !string.IsNullOrEmpty(Settings.Default.UserName) && Settings.Default.IsDirty;
        }

        private void UpdateWorkerStatus()
        {
            // disable (or enable) any action items if the worker is busy (or not)
            ResetMatches.Enabled = !Worker.IsBusy;
            SyncMenuItem.Enabled = !Worker.IsBusy;
            ExitMenuItem.Enabled = !Worker.IsBusy;
            WorkTimer.Enabled = Worker.IsBusy;

            // reset the notification icon if the worker has finished
            if (!Worker.IsBusy)
            {
                WorkTimer.Tag = 0;
                Notifications.Icon = Resources.Idle;
            }
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // if the user closed the form, reset its content and hide it but don't close the app itself (same goes if we're still syncing)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Settings.Default.Reload();
                Hide();
            }
            else if (Worker.IsBusy)
                e.Cancel = true;
        }

        private void SettingsForm_Shown(object sender, EventArgs e)
        {
            // hide the window if the user name was already entered
            if (!string.IsNullOrEmpty(Settings.Default.UserName))
                Hide();
        }

        private void Settings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // handle relevant settings and update the save status
            var settings = (Settings)sender;
            switch (e.PropertyName)
            {
                case "UserName":
                    // if the user name was changed (and the settings are dirty, ie. not reloaded etc) then reset the last sync time
                    if (settings.IsDirty)
                        settings.IsFirstSync = true;
                    break;
                case "IsDirty":
                    break;
                default:
                    return;
            }
            UpdateSaveStatus();
        }

        private void SyncTimer_Tick(object sender, EventArgs e)
        {
            // resync if the necessary amount of time has elapsed
            if (DateTime.Now - Settings.Default.LastSync > Settings.Default.SyncInterval)
                Sync(false, false);
        }

        private void ResetMatches_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // intiate reset matches
            Sync(true, true);
        }

        private void GoogleContactsSignup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // show the Google contacts page (if the user already signed up, this will remind him or her of that fact :)
            Process.Start("https://www.google.com/contacts/");
        }

        private void Help_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // show the GO Contact Sync Mod page (yeah, it's a rename nightmare)
            Process.Start("http://googlesyncmod.sourceforge.net/");
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // retrieve the last report and update the notification status
            var context = (SyncContext)e.UserState;
            string text;
            ToolTipIcon icon;
            context.GetLastReport(out text, out icon);
            UpdateNotificationStatus(text, icon, context.Interactive);
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // log the start
            Logger.Log("Worker started.", EventType.Debug);

            // create vars
            var worker = (BackgroundWorker)sender;
            var context = (SyncContext)e.Argument;
            e.Result = context;
            string duplicates = null;

            // initialize the syncer
            context.Report(worker, Resources.SettingsForm_InitializeSync, ToolTipIcon.Info);
            var sync = new Syncronizer()
            {
                SyncProfile = WindowsIdentity.GetCurrent().User.Value,
                SyncOption = context.Mode,
                SyncDelete = true,
                PromptDelete = false,
                UseFileAs = true,
                SyncNotes = false,
                SyncContacts = true,
            };
            sync.ErrorEncountered += (title, ex, type) =>
            {
                // log the error and report it through the worker
                Logger.Log(ex.Message, type);
                ToolTipIcon icon;
                switch (type)
                {
                    case EventType.Information: icon = ToolTipIcon.Info; break;
                    case EventType.Error: icon = ToolTipIcon.Error; break;
                    case EventType.Warning: icon = ToolTipIcon.Warning; break;
                    default: icon = ToolTipIcon.None; break;
                }
                context.Report(worker, ex.Message, icon, true);
            };
            sync.DuplicatesFound += (title, text) => duplicates = text;

            // log into Google
            context.Report(worker, Resources.SettingsForm_GoogleLogon, ToolTipIcon.Info);
            sync.LoginToGoogle(context.UserName, DecodePassword(context.Password));
            try
            {
                // access Outlook
                context.Report(worker, Resources.SettingsForm_OutlookLogon, ToolTipIcon.Info);
                sync.LoginToOutlook();
                try
                {
                    // set the proper folder (this is necessary since some parts of ContactsMatcher don't check for null or empty)
                    Syncronizer.SyncContactsFolder = Syncronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts).EntryID;

                    // reset matches
                    if ((context.Tasks & WorkerTasks.ResetMatches) != 0)
                    {
                        context.Report(worker, Resources.SettingsForm_ResetMatches, ToolTipIcon.Info);
                        sync.LoadContacts();
                        sync.ResetContactMatches();
                    }

                    // sync
                    if ((context.Tasks & WorkerTasks.Synchronize) != 0)
                    {
                        context.Report(worker, Resources.SettingsForm_SyncContacts, ToolTipIcon.Info);
                        sync.Sync();
                    }
                }
                finally
                {
                    // log out from Outlook
                    context.Report(worker, Resources.SettingsForm_OutlookLogoff, ToolTipIcon.Info);
                    sync.LogoffOutlook();
                }
            }
            finally
            {
                // log out from Google
                context.Report(worker, Resources.SettingsForm_GoogleLogoff, ToolTipIcon.Info);
                sync.LogoffGoogle();
            }

            // finalizing
            context.Report
            (
                worker,
                string.Format(string.IsNullOrEmpty(duplicates) ? Resources.SettingsForm_SyncResult : Resources.SettingsForm_SyncResultWithDuplicates, DateTime.Now, sync.TotalCount, sync.SyncedCount, sync.DeletedCount, sync.SkippedCount, sync.ErrorCount, duplicates),
                sync.ErrorCount > 0 ? ToolTipIcon.Error : (sync.SkippedCount > 0 || !string.IsNullOrEmpty(duplicates)) ? ToolTipIcon.Warning : ToolTipIcon.Info,
                true
            );

            // log the result
            Logger.Log("Worker ended.", EventType.Debug);
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var context = e.Error == null ? (SyncContext)e.Result : null;

            // reflect the new worker status
            UpdateWorkerStatus();

            // update the settings (set the last sync time and reset the first sync flag - before displaying any dialog boxes!)
            var wasDirty = Settings.Default.IsDirty;
            Settings.Default.LastSync = DateTime.Now;
            if (context != null && context.UserName == Settings.Default.UserName)
                Settings.Default.IsFirstSync = false;
            if (!wasDirty)
                Settings.Default.Save();

            // log and display the error to the user
            if (context == null)
            {
                Logger.Log(e.Error.ToString(), EventType.Error);
                UpdateNotificationStatus(e.Error.Message, ToolTipIcon.Error, true);
                MessageBox.Show(e.Error.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            // hide the form and reset its content
            Hide();
            Settings.Default.Reload();
        }

        private void Save_Click(object sender, EventArgs e)
        {
            // hide the form, save the settings and start a new sync
            Hide();
            Settings.Default.Save();
            Sync(false, true);
        }

        private void SyncMenuItem_Click(object sender, EventArgs e)
        {
            // start a new sync
            Sync(false, true);
        }

        private void OptionsMenuItem_Click(object sender, EventArgs e)
        {
            // show the form and activate it
            Show();
            Activate();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            // exit the application
            Application.Exit();
        }

        private void WorkTimer_Tick(object sender, EventArgs e)
        {
            // animate the notification icon
            Notifications.Icon = WorkIcons[(int)WorkTimer.Tag];
            WorkTimer.Tag = ((int)WorkTimer.Tag + 1) % WorkIcons.Length;
        }

        private void Notifications_MouseClick(object sender, MouseEventArgs e)
        {
            // start the timer to differentiate between single and double click
            if (e.Button == MouseButtons.Left && !DoubleClickTimer.Enabled)
            {
                DoubleClickTimer.Interval = SystemInformation.DoubleClickTime;
                DoubleClickTimer.Enabled = true;
            }
        }

        private void Notifications_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // stop the timer and call the options action
            if (e.Button == MouseButtons.Left && DoubleClickTimer.Enabled)
            {
                DoubleClickTimer.Enabled = false;
                OptionsMenuItem_Click(sender, e);
            }
        }

        private void DoubleClickTimer_Tick(object sender, EventArgs e)
        {
            // stop the timer and show the balloon, if this isn't a lingering tick
            if (DoubleClickTimer.Enabled)
            {
                DoubleClickTimer.Enabled = false;
                if (!string.IsNullOrEmpty(Notifications.BalloonTipText))
                    Notifications.ShowBalloonTip((int)BalloonTimeout.TotalMilliseconds);
            }
        }
    }

    namespace Properties
    {
        internal sealed partial class Settings
        {
            private bool isDirty = false;

            protected override void OnSettingChanging(object sender, System.Configuration.SettingChangingEventArgs e)
            {
                // set the dirty state if the change hasn't been cancelled
                base.OnSettingChanging(sender, e);
                if (!e.Cancel)
                    this.IsDirty = true;
            }

            protected override void OnSettingsSaving(object sender, CancelEventArgs e)
            {
                // reset the dirty flag if the save operation hasn't been cancelled
                base.OnSettingsSaving(sender, e);
                if (!e.Cancel)
                    this.IsDirty = false;
            }

            public override object this[string propertyName]
            {
                get
                {
                    // return the current value
                    return base[propertyName];
                }
                set
                {
                    // compare the values and set the new one if it's different
                    if (!object.Equals(value, base[propertyName]))
                        base[propertyName] = value;
                }
            }

            public new void Reset()
            {
                this.IsDirty = false;
                base.Reset();
            }

            public new void Reload()
            {
                this.IsDirty = false;
                base.Reload();
            }

            public new void Upgrade()
            {
                this.IsDirty = false;
                base.Upgrade();
            }

            public bool IsDirty
            {
                get { return this.isDirty; }
                private set
                {
                    // update the flag if it has changed and notify any listeners
                    if (value == this.isDirty)
                        return;
                    this.isDirty = value;
                    this.OnPropertyChanged(this, new PropertyChangedEventArgs("IsDirty"));
                }
            }
        }
    }
}
