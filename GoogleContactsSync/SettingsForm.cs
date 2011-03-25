using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using System.Runtime.Remoting;

namespace GoContactSyncMod
{
	internal partial class SettingsForm : Form
	{
		internal Syncronizer _sync;
		private SyncOption _syncOption;
		private DateTime lastSync;
		private bool requestClose = false;

		delegate void TextHandler(string text);
		delegate void SwitchHandler(bool value);

		public SettingsForm()
		{
			InitializeComponent();
			Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);
			PopulateSyncOptionBox();

			LoadSettings();

			lastSync = DateTime.Now.AddSeconds(15) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
			lastSyncLabel.Text = "Not synced";

			ValidateSyncButton();
		}

		private void PopulateSyncOptionBox()
		{
			string str;
			for (int i = 0; i < 20; i++)
			{
				str = ((SyncOption)i).ToString();
				if (str == i.ToString())
					break;

				// format (to add space before capital)
				MatchCollection matches = Regex.Matches(str, "[A-Z]");
				for (int k = 0; k < matches.Count; k++)
				{
					str = str.Replace(str[matches[k].Index].ToString(), " " + str[matches[k].Index]);
					matches = Regex.Matches(str, "[A-Z]");
				}
				str = str.Replace("  ", " ");
				// fix start
				str = str.Substring(1);

				syncOptionBox.Items.Add(str);
			}
		}

		private void LoadSettings()
		{
			// default
			SetSyncOption(0);

			// load
			RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
			if (regKeyAppRoot.GetValue("SyncOption") != null)
			{
				_syncOption = (SyncOption)regKeyAppRoot.GetValue("SyncOption");
				SetSyncOption((int)_syncOption);
			}
			if (regKeyAppRoot.GetValue("SyncProfile") != null)
				tbSyncProfile.Text = (string)regKeyAppRoot.GetValue("SyncProfile");
			if (regKeyAppRoot.GetValue("Username") != null)
			{
				UserName.Text = regKeyAppRoot.GetValue("Username") as string;
				if (regKeyAppRoot.GetValue("Password") != null)
					Password.Text = Encryption.DecryptPassword(UserName.Text, regKeyAppRoot.GetValue("Password") as string);
			}
			if (regKeyAppRoot.GetValue("AutoSync") != null)
				autoSyncCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("AutoSync"));
			if (regKeyAppRoot.GetValue("AutoSyncInterval") != null)
				autoSyncInterval.Value = Convert.ToDecimal(regKeyAppRoot.GetValue("AutoSyncInterval"));
			if (regKeyAppRoot.GetValue("AutoStart") != null)
				runAtStartupCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("AutoStart"));
			if (regKeyAppRoot.GetValue("ReportSyncResult") != null)
				reportSyncResultCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("ReportSyncResult"));
			if (regKeyAppRoot.GetValue("SyncDeletion") != null)
				btSyncDelete.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncDeletion"));

			autoSyncCheckBox_CheckedChanged(null, null);
		}
		private void SaveSettings()
		{
			RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
			regKeyAppRoot.SetValue("SyncOption", (int)_syncOption);
			if (!string.IsNullOrEmpty(tbSyncProfile.Text))
				regKeyAppRoot.SetValue("SyncProfile", tbSyncProfile.Text);
			if (!string.IsNullOrEmpty(UserName.Text))
			{
				regKeyAppRoot.SetValue("Username", UserName.Text);
				if (!string.IsNullOrEmpty(Password.Text))
					regKeyAppRoot.SetValue("Password", Encryption.EncryptPassword(UserName.Text, Password.Text));
			}
			regKeyAppRoot.SetValue("AutoSync", autoSyncCheckBox.Checked.ToString());
			regKeyAppRoot.SetValue("AutoSyncInterval", autoSyncInterval.Value.ToString());
			regKeyAppRoot.SetValue("AutoStart", runAtStartupCheckBox.Checked);
			regKeyAppRoot.SetValue("ReportSyncResult", reportSyncResultCheckBox.Checked);
			regKeyAppRoot.SetValue("SyncDeletion", btSyncDelete.Checked);
		}

		private bool ValidCredentials
		{
			get
			{
				bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
				bool passwordIsValid = Password.Text.Length != 0;
				bool syncProfileNameIsValid = tbSyncProfile.Text.Length != 0;

				setBgColor(UserName, userNameIsValid);
				setBgColor(Password, passwordIsValid);
				setBgColor(tbSyncProfile, syncProfileNameIsValid);
				return userNameIsValid && passwordIsValid && syncProfileNameIsValid;
			}
		}
		private void setBgColor(TextBox box, bool isValid)
		{
			if (!isValid)
				box.BackColor = Color.LightPink;
			else
				box.BackColor = Color.LightGreen;
		}

		private void button4_Click(object sender, EventArgs e)
		{
			Sync();
		}
		private void Sync()
		{
			try
			{
				if (!ValidCredentials)
					return;

				//Sync_ThreadStarter();

				ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
				Thread thread = new Thread(starter);
				thread.Start();

				// wait for thread to start
				while (!thread.IsAlive)
					Thread.Sleep(1);
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

		private void Sync_ThreadStarter()
		{
			try
			{
				TimerSwitch(false);
				SetLastSyncText("Syncing...");
                notifyIcon.Text = Application.ProductName + "\nSyncing...";
				SetFormEnabled(false);

				if (_sync == null)
				{
					_sync = new Syncronizer();
					_sync.DuplicatesFound += new Syncronizer.NotificationHandler(_sync_DuplicatesFound);
					_sync.ErrorEncountered += new Syncronizer.ErrorNotificationHandler(_sync_ErrorEncountered);
				}

				Logger.ClearLog();
				SetSyncConsoleText("");
				Logger.Log("Sync started.", EventType.Information);
				//SetSyncConsoleText(Logger.GetText());
				_sync.SyncProfile = tbSyncProfile.Text;
				_sync.SyncOption = _syncOption;

                try
                {
                    _sync.LoginToGoogle(UserName.Text, Password.Text);
                    _sync.LoginToOutlook();

                    _sync.Sync();
                    
                    lastSync = DateTime.Now;
                    SetLastSyncText("Last synced at " + lastSync.ToString());

                    string message = string.Format("Sync complete.\r\n Synced:  {1} out of {0}.\r\n Deleted:  {2}.\r\n Skipped: {3}.\r\n Errors:    {4}.", _sync.TotalCount, _sync.SyncedCount, _sync.DeletedCount, _sync.SkippedCount, _sync.ErrorCount);
                    Logger.Log(message, EventType.Information);
                    if (reportSyncResultCheckBox.Checked)
                    {
                        notifyIcon.BalloonTipTitle = Application.ProductName;
                        notifyIcon.BalloonTipText = string.Format("{0}. {1}", DateTime.Now, message);

                        if (_sync.ErrorCount > 0)
                            notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                        else if (_sync.SkippedCount > 0)
                            notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
                        else
                            notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                        notifyIcon.ShowBalloonTip(5000);
                    }
                    string toolTip = string.Format("{0}\nLast sync: {1}", Application.ProductName, DateTime.Now.ToString("dd.MM. HH:mm"));
                    if (_sync.ErrorCount + _sync.SkippedCount > 0)
                        toolTip += string.Format("\nWarnings: {0}.", _sync.ErrorCount + _sync.SkippedCount);
                    if (toolTip.Length >= 64)
                        toolTip = toolTip.Substring(0, 63);
                    notifyIcon.Text = toolTip;
                }
                catch (Google.GData.Client.GDataRequestException ex)
                {
                    SetLastSyncText("Sync failed.");
                    notifyIcon.Text = Application.ProductName + "\nSync failed";

                    string responseString = ex.ResponseString;
                    if (ex.InnerException is System.Net.WebException)
                    {
                        Logger.Log("Cannot connect to Google, please check for available internet connection and proxy settings if applicable: " + ((System.Net.WebException)ex.InnerException).Message + "\r\n" + responseString, EventType.Warning);
                    }
                    else
                    {
                        ErrorHandler.Handle(ex);
                    }
                }
                catch (Exception ex)
                {
                    SetLastSyncText("Sync failed.");
                    notifyIcon.Text = Application.ProductName + "\nSync failed";
                    ErrorHandler.Handle(ex);
                }				

			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
			finally
			{                        
                lastSync = DateTime.Now;
                TimerSwitch(true);
				SetFormEnabled(true);
                if (_sync != null)
                {
                    _sync.LogoffGoogle();
                    _sync = null;
                }
			}
		}

		void Logger_LogUpdated(string Message)
		{
			AppendSyncConsoleText(Message);
		}

		void _sync_ErrorEncountered(string title, Exception ex, EventType eventType)
		{
			// do not show ErrorHandler, as there may be multiple exceptions that would nag the user
			Logger.Log(ex.ToString(), EventType.Error);
			string message = String.Format("Error Saving Contact: {0}.\nPlease report complete ErrorMessage from Log to the Tracker\nat https://sourceforge.net/tracker/?group_id=369321", ex.Message);
			notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
			notifyIcon.ShowBalloonTip(5000);
		}

		void _sync_DuplicatesFound(string title, string message, EventType eventType)
		{
			Logger.Log(message, EventType.Warning);
			notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
			notifyIcon.ShowBalloonTip(5000);
		}

		public void SetFormEnabled(bool enabled)
		{
			if (this.InvokeRequired)
			{
				SwitchHandler h = new SwitchHandler(SetFormEnabled);
				this.Invoke(h, new object[] { enabled });
			}
			else
			{
				resetMatchesLinkLabel.Enabled = enabled;
				settingsGroupBox.Enabled = enabled;
				syncButton.Enabled = enabled;
			}
		}
		public void SetLastSyncText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(SetLastSyncText);
				this.Invoke(h, new object[] { text });
			}
			else
				lastSyncLabel.Text = text;
		}
		public void SetSyncConsoleText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(SetSyncConsoleText);
				this.Invoke(h, new object[] { text });
			}
			else
            {
				syncConsole.Text = text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }

		}
		public void AppendSyncConsoleText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(AppendSyncConsoleText);
				this.Invoke(h, new object[] { text });
			}
			else
            {
				syncConsole.Text += text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }
		}
		public void TimerSwitch(bool value)
		{
			if (this.InvokeRequired)
			{
				SwitchHandler h = new SwitchHandler(TimerSwitch);
				this.Invoke(h, new object[] { value });
			}
			else
			{

				if (value)
				{
					if (autoSyncCheckBox.Checked)
					{
						autoSyncInterval.Enabled = autoSyncCheckBox.Checked;
						syncTimer.Enabled = autoSyncCheckBox.Checked;
						nextSyncLabel.Visible = autoSyncCheckBox.Checked;
					}
				}
				else
				{                    
					autoSyncInterval.Enabled = value;
                    syncTimer.Enabled = value;
					nextSyncLabel.Visible = value;
				}
			}
		}

		// Fix for WinXP and older systems, that do not continue with shutdown until all programs have closed
		// FormClosing would hold system shutdown, when it sets the cancel to true
		private static int WM_QUERYENDSESSION = 0x11;
		protected override void WndProc(ref System.Windows.Forms.Message m)
		{
			if (m.Msg == WM_QUERYENDSESSION)
				requestClose = true;

			// If this is WM_QUERYENDSESSION, the form must exit and not just hide
			base.WndProc(ref m);
		} 
		private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (!requestClose)
			{
				SaveSettings();
				e.Cancel = true;
			}
			HideForm();
		}
		private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
		{
			try
			{
				if (_sync != null)
					_sync.LogoffOutlook();

				SaveSettings();

				notifyIcon.Dispose();
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

		private void syncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				int index = syncOptionBox.SelectedIndex;
				if (index == -1)
					return;

				SetSyncOption(index);
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}
		private void SetSyncOption(int index)
		{
			_syncOption = (SyncOption)index;
			for (int i = 0; i < syncOptionBox.Items.Count; i++)
			{
				if (i == index)
					syncOptionBox.SetItemCheckState(i, CheckState.Checked);
				else
					syncOptionBox.SetItemCheckState(i, CheckState.Unchecked);
			}
		}

		private void SettingsForm_Resize(object sender, EventArgs e)
		{
			if (WindowState == FormWindowState.Minimized)
				Hide();

		}

		private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
		{
			if (WindowState == FormWindowState.Normal)
				HideForm();
			else
				ShowForm();
		}

		private void autoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
		{
            lastSync = DateTime.Now.AddSeconds(15) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
			autoSyncInterval.Enabled = autoSyncCheckBox.Checked;
			syncTimer.Enabled = autoSyncCheckBox.Checked;
			nextSyncLabel.Visible = autoSyncCheckBox.Checked;
		}

		private void syncTimer_Tick(object sender, EventArgs e)
		{
			if (lastSync != null)
			{
				TimeSpan syncTime = DateTime.Now - lastSync;
				TimeSpan limit = new TimeSpan(0, (int)autoSyncInterval.Value, 0);
				if (syncTime < limit)
				{
					TimeSpan diff = limit - syncTime;
					string str = "Next sync in";
					if (diff.Hours != 0)
						str += " " + diff.Hours + " h";
					if (diff.Minutes != 0 || diff.Hours != 0)
						str += " " + diff.Minutes + " min";
					if (diff.Seconds != 0)
						str += " " + diff.Seconds + " s";
					nextSyncLabel.Text = str;
					return;
				}
			}
			Sync();
		}

		private void resetMatchesLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

			// force deactivation to show up
			Application.DoEvents();
			try
			{
                TimerSwitch(false);
                SetLastSyncText("Resetting matches...");
                notifyIcon.Text = Application.ProductName + "\nResetting matches...";
                SetFormEnabled(false);
                this.hideButton.Enabled = false;

				if (_sync == null)
				{
					_sync = new Syncronizer();
				}

                Logger.ClearLog();
                SetSyncConsoleText("");
                Logger.Log("Reset Matches started.", EventType.Information);

				_sync.LoginToGoogle(UserName.Text, Password.Text);
				_sync.LoginToOutlook();
                _sync.SyncProfile = tbSyncProfile.Text;

                //Load matches, but match them by properties, not sync id
				_sync.Load();

				_sync.ResetMatches();

                lastSync = DateTime.Now;
                SetLastSyncText("Matches reset at " + lastSync.ToString());
                Logger.Log("Matches reset.", EventType.Information);                
			}
			catch (Exception ex)
            {
                SetLastSyncText("Reset Matches failed");
                Logger.Log("Reset Matches failed", EventType.Error);
				ErrorHandler.Handle(ex);
			}
			finally
			{
                lastSync = DateTime.Now;
                TimerSwitch(true);
				SetFormEnabled(true);
                this.hideButton.Enabled = true;
			}
		}

        private void ShowForm()
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
		private void HideForm()
		{
			WindowState = FormWindowState.Minimized;
			Hide();
		}

		private void toolStripMenuItem1_Click(object sender, EventArgs e)
		{
			ShowForm();
            this.Activate();
		}
		private void toolStripMenuItem3_Click(object sender, EventArgs e)
		{
			HideForm();
		}
		private void toolStripMenuItem2_Click(object sender, EventArgs e)
		{
			requestClose = true;
			Close();
		}
		private void toolStripMenuItem5_Click(object sender, EventArgs e)
		{
			AboutBox about = new AboutBox();
			about.Show();
		}
		private void toolStripMenuItem4_Click(object sender, EventArgs e)
		{
			Sync();
		}

		private void SettingsForm_Load(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(UserName.Text) ||
				string.IsNullOrEmpty(Password.Text) ||
				string.IsNullOrEmpty(tbSyncProfile.Text))
			{
				// this is the first load, show form
				ShowForm();
				UserName.Focus();
			}
			else
				HideForm();
		}

		private void runAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
		{
			RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");

			if (runAtStartupCheckBox.Checked)
			{
				// add to registry
				regKeyAppRoot.SetValue("GoogleContactSync", Application.ExecutablePath);
			}
			else
			{
				// remove from registry
				regKeyAppRoot.DeleteValue("GoogleContactSync");
			}
		}

		private void UserName_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}
		private void Password_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}

		private void ValidateSyncButton()
		{
			syncButton.Enabled = ValidCredentials;
		}

		private void deleteDuplicatesButton_Click(object sender, EventArgs e)
		{
			//DeleteDuplicatesForm f = new DeleteDuplicatesForm(_sync
		}

		private void tbSyncProfile_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}

		private void btSyncDelete_CheckedChanged(object sender, EventArgs e)
		{
			if (_sync == null)
			{
				_sync = new Syncronizer();
			}

			_sync.SyncDelete = btSyncDelete.Checked;
		}

		private void Donate_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start("https://sourceforge.net/project/project_donations.php?group_id=369321");
		}

		private void Donate_MouseEnter(object sender, EventArgs e)
		{
			Donate.BackColor = System.Drawing.Color.LightGray;
		}

		private void Donate_MouseLeave(object sender, EventArgs e)
		{
			Donate.BackColor = System.Drawing.Color.Transparent;
		}

		private void hideButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void proxySettingsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			// TODO: Implement a dialog with proxy settings and use them when connecting with Google

			// Alpha quick'm'dirty workaround solution
			try
			{
				if (MessageBox.Show("The proxy configuration is in beta stage, a more comfortable solution is to come. For now, you have to edit the Applications Config file with administrator privileges.\n\nOpen Configuration file now?",
					"GO Contact Sync Mod", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.OK)
				{
					StartFileSystemWatcher();
					try
					{
						ProcessStartInfo psi = new ProcessStartInfo();
						psi.Arguments = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
						// full path not required, notepad is usually installed in system directory
						psi.FileName = "notepad.exe";
						// Vista and Win7 UAC Control
						psi.Verb = "runas";
						Process.Start(psi);
					}
					catch (Exception)
					{
						// fallback if notepad is not found/installed
						// in default configuration this will open Internet Explorer, but user can then see the path and open an editor himself
						Process.Start(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
					}
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

		FileSystemWatcher fsw = null;
		private void StartFileSystemWatcher()
		{
			if (fsw == null)
			{
				fsw = new FileSystemWatcher();
				fsw.Path = Path.GetDirectoryName(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
				fsw.Filter = Path.GetFileName(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
				fsw.NotifyFilter = NotifyFilters.LastWrite;
				fsw.Changed += delegate
				{
					this.BeginInvoke((MethodInvoker)delegate
					{
						MessageBox.Show(this, "If you have changed your proxy configuration, restart the Program to take effect.", "GO Contact Sync Mod");
					});
					// cleanup after we've notified the user
					fsw.EnableRaisingEvents = false;
					fsw.Dispose();
					fsw = null;
				};
				fsw.EnableRaisingEvents = true;
			}
		}

		private void SettingsForm_HelpButtonClicked(object sender, CancelEventArgs e)
		{
			ShowHelp();
		}

		private void SettingsForm_HelpRequested(object sender, HelpEventArgs hlpevent)
		{
			ShowHelp();
		}

		private void ShowHelp()
		{
			// go to the page showing the help and howto instructions
			Process.Start("http://googlesyncmod.sourceforge.net/");
		}
     

	}

	//internal class EventLogger : ILogger
	//{
	//    private TextBox _box;
	//    public EventLogger(TextBox box)
	//    {
	//        _box = box;
	//    }

	//    #region ILogger Members

	//    public void Log(string message, EventType eventType)
	//    {
	//        _box.Text += message;
	//    }

	//    #endregion
	//}
}