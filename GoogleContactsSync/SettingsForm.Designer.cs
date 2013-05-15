namespace GoContactSyncMod
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.Password = new System.Windows.Forms.TextBox();
            this.PasswordLabel = new System.Windows.Forms.Label();
            this.UserNameLabel = new System.Windows.Forms.Label();
            this.Notifications = new System.Windows.Forms.NotifyIcon(this.components);
            this.NotificationMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.SyncMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.FirstMenuSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.OptionsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SecondMenuSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.ExitMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SyncTimer = new System.Windows.Forms.Timer(this.components);
            this.ResetMatches = new System.Windows.Forms.LinkLabel();
            this.ContactsImage = new System.Windows.Forms.PictureBox();
            this.SyncInterval = new System.Windows.Forms.NumericUpDown();
            this.SyncIntervalUnit = new System.Windows.Forms.Label();
            this.SyncIntervalLabel = new System.Windows.Forms.Label();
            this.TwoWaySync = new System.Windows.Forms.RadioButton();
            this.AccountHeader = new System.Windows.Forms.Label();
            this.GoogleContactsImage = new System.Windows.Forms.PictureBox();
            this.Description1 = new System.Windows.Forms.Label();
            this.Description2 = new System.Windows.Forms.Label();
            this.OptionsHeader = new System.Windows.Forms.Label();
            this.TwoWaySyncDescription1 = new System.Windows.Forms.Label();
            this.TwoWaySyncDescription2 = new System.Windows.Forms.Label();
            this.GoogleToOutlook = new System.Windows.Forms.RadioButton();
            this.GoogleToOutlookDescription = new System.Windows.Forms.Label();
            this.OutlookToGoogle = new System.Windows.Forms.RadioButton();
            this.OutlookToGoogleDescription = new System.Windows.Forms.Label();
            this.Save = new System.Windows.Forms.Button();
            this.GoogleContactsSignup = new System.Windows.Forms.LinkLabel();
            this.Cancel = new System.Windows.Forms.Button();
            this.Help = new System.Windows.Forms.LinkLabel();
            this.UserName = new System.Windows.Forms.TextBox();
            this.Worker = new System.ComponentModel.BackgroundWorker();
            this.WorkTimer = new System.Windows.Forms.Timer(this.components);
            this.DoubleClickTimer = new System.Windows.Forms.Timer(this.components);
            this.NotificationMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ContactsImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.SyncInterval)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GoogleContactsImage)).BeginInit();
            this.SuspendLayout();
            // 
            // Password
            // 
            this.Password.Location = new System.Drawing.Point(109, 168);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.Size = new System.Drawing.Size(224, 22);
            this.Password.TabIndex = 6;
            this.Password.UseSystemPasswordChar = true;
            // 
            // PasswordLabel
            // 
            this.PasswordLabel.AutoSize = true;
            this.PasswordLabel.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PasswordLabel.Location = new System.Drawing.Point(23, 174);
            this.PasswordLabel.Name = "PasswordLabel";
            this.PasswordLabel.Size = new System.Drawing.Size(73, 16);
            this.PasswordLabel.TabIndex = 5;
            this.PasswordLabel.Text = "Password:";
            // 
            // UserNameLabel
            // 
            this.UserNameLabel.AutoSize = true;
            this.UserNameLabel.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UserNameLabel.Location = new System.Drawing.Point(48, 138);
            this.UserNameLabel.Name = "UserNameLabel";
            this.UserNameLabel.Size = new System.Drawing.Size(48, 16);
            this.UserNameLabel.TabIndex = 3;
            this.UserNameLabel.Text = "Email:";
            // 
            // Notifications
            // 
            this.Notifications.ContextMenuStrip = this.NotificationMenu;
            this.Notifications.Visible = true;
            this.Notifications.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Notifications_MouseClick);
            this.Notifications.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Notifications_MouseDoubleClick);
            // 
            // NotificationMenu
            // 
            this.NotificationMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.SyncMenuItem,
            this.FirstMenuSeparator,
            this.OptionsMenuItem,
            this.SecondMenuSeparator,
            this.ExitMenuItem});
            this.NotificationMenu.Name = "systemTrayMenu";
            this.NotificationMenu.Size = new System.Drawing.Size(112, 82);
            // 
            // SyncMenuItem
            // 
            this.SyncMenuItem.Name = "SyncMenuItem";
            this.SyncMenuItem.Size = new System.Drawing.Size(111, 22);
            this.SyncMenuItem.Text = "Sync";
            this.SyncMenuItem.Click += new System.EventHandler(this.SyncMenuItem_Click);
            // 
            // FirstMenuSeparator
            // 
            this.FirstMenuSeparator.Name = "FirstMenuSeparator";
            this.FirstMenuSeparator.Size = new System.Drawing.Size(108, 6);
            // 
            // OptionsMenuItem
            // 
            this.OptionsMenuItem.Name = "OptionsMenuItem";
            this.OptionsMenuItem.Size = new System.Drawing.Size(111, 22);
            this.OptionsMenuItem.Text = "Options";
            this.OptionsMenuItem.Click += new System.EventHandler(this.OptionsMenuItem_Click);
            // 
            // SecondMenuSeparator
            // 
            this.SecondMenuSeparator.Name = "SecondMenuSeparator";
            this.SecondMenuSeparator.Size = new System.Drawing.Size(108, 6);
            // 
            // ExitMenuItem
            // 
            this.ExitMenuItem.Name = "ExitMenuItem";
            this.ExitMenuItem.Size = new System.Drawing.Size(111, 22);
            this.ExitMenuItem.Text = "Exit";
            this.ExitMenuItem.Click += new System.EventHandler(this.ExitMenuItem_Click);
            // 
            // SyncTimer
            // 
            this.SyncTimer.Enabled = true;
            this.SyncTimer.Interval = 60000;
            this.SyncTimer.Tick += new System.EventHandler(this.SyncTimer_Tick);
            // 
            // ResetMatches
            // 
            this.ResetMatches.AutoSize = true;
            this.ResetMatches.Location = new System.Drawing.Point(106, 460);
            this.ResetMatches.Name = "ResetMatches";
            this.ResetMatches.Size = new System.Drawing.Size(96, 16);
            this.ResetMatches.TabIndex = 19;
            this.ResetMatches.TabStop = true;
            this.ResetMatches.Text = "Reset Matches";
            this.ResetMatches.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ResetMatches_LinkClicked);
            // 
            // ContactsImage
            // 
            this.ContactsImage.Image = ((System.Drawing.Image)(resources.GetObject("ContactsImage.Image")));
            this.ContactsImage.Location = new System.Drawing.Point(326, 103);
            this.ContactsImage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ContactsImage.Name = "ContactsImage";
            this.ContactsImage.Size = new System.Drawing.Size(128, 128);
            this.ContactsImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.ContactsImage.TabIndex = 6;
            this.ContactsImage.TabStop = false;
            // 
            // SyncInterval
            // 
            this.SyncInterval.Location = new System.Drawing.Point(189, 406);
            this.SyncInterval.Maximum = new decimal(new int[] {
            1440,
            0,
            0,
            0});
            this.SyncInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.SyncInterval.Name = "SyncInterval";
            this.SyncInterval.Size = new System.Drawing.Size(51, 22);
            this.SyncInterval.TabIndex = 17;
            this.SyncInterval.Value = new decimal(new int[] {
            120,
            0,
            0,
            0});
            // 
            // SyncIntervalUnit
            // 
            this.SyncIntervalUnit.AutoSize = true;
            this.SyncIntervalUnit.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SyncIntervalUnit.Location = new System.Drawing.Point(251, 408);
            this.SyncIntervalUnit.Name = "SyncIntervalUnit";
            this.SyncIntervalUnit.Size = new System.Drawing.Size(58, 16);
            this.SyncIntervalUnit.TabIndex = 18;
            this.SyncIntervalUnit.Text = "minutes";
            // 
            // SyncIntervalLabel
            // 
            this.SyncIntervalLabel.AutoSize = true;
            this.SyncIntervalLabel.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SyncIntervalLabel.Location = new System.Drawing.Point(106, 408);
            this.SyncIntervalLabel.Name = "SyncIntervalLabel";
            this.SyncIntervalLabel.Size = new System.Drawing.Size(78, 16);
            this.SyncIntervalLabel.TabIndex = 16;
            this.SyncIntervalLabel.Text = "Sync every";
            // 
            // TwoWaySync
            // 
            this.TwoWaySync.AutoSize = true;
            this.TwoWaySync.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TwoWaySync.Location = new System.Drawing.Point(90, 247);
            this.TwoWaySync.Name = "TwoWaySync";
            this.TwoWaySync.Size = new System.Drawing.Size(63, 20);
            this.TwoWaySync.TabIndex = 9;
            this.TwoWaySync.Text = "2-way";
            this.TwoWaySync.UseVisualStyleBackColor = true;
            // 
            // AccountHeader
            // 
            this.AccountHeader.AutoSize = true;
            this.AccountHeader.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AccountHeader.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(23)))), ((int)(((byte)(71)))), ((int)(((byte)(171)))));
            this.AccountHeader.Location = new System.Drawing.Point(8, 90);
            this.AccountHeader.Name = "AccountHeader";
            this.AccountHeader.Size = new System.Drawing.Size(198, 19);
            this.AccountHeader.TabIndex = 2;
            this.AccountHeader.Text = "Google Account Settings";
            // 
            // GoogleContactsImage
            // 
            this.GoogleContactsImage.Image = ((System.Drawing.Image)(resources.GetObject("GoogleContactsImage.Image")));
            this.GoogleContactsImage.Location = new System.Drawing.Point(5, 11);
            this.GoogleContactsImage.Name = "GoogleContactsImage";
            this.GoogleContactsImage.Size = new System.Drawing.Size(170, 71);
            this.GoogleContactsImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.GoogleContactsImage.TabIndex = 13;
            this.GoogleContactsImage.TabStop = false;
            // 
            // Description1
            // 
            this.Description1.AutoSize = true;
            this.Description1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Description1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(95)))), ((int)(((byte)(116)))), ((int)(((byte)(138)))));
            this.Description1.Location = new System.Drawing.Point(216, 24);
            this.Description1.Name = "Description1";
            this.Description1.Size = new System.Drawing.Size(158, 15);
            this.Description1.TabIndex = 0;
            this.Description1.Text = "Sync Google Contacts with";
            // 
            // Description2
            // 
            this.Description2.AutoSize = true;
            this.Description2.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Description2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(95)))), ((int)(((byte)(116)))), ((int)(((byte)(138)))));
            this.Description2.Location = new System.Drawing.Point(216, 40);
            this.Description2.Name = "Description2";
            this.Description2.Size = new System.Drawing.Size(173, 15);
            this.Description2.TabIndex = 1;
            this.Description2.Text = "Microsoft Outlook™ contacts";
            // 
            // OptionsHeader
            // 
            this.OptionsHeader.AutoSize = true;
            this.OptionsHeader.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OptionsHeader.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(23)))), ((int)(((byte)(71)))), ((int)(((byte)(171)))));
            this.OptionsHeader.Location = new System.Drawing.Point(8, 218);
            this.OptionsHeader.Name = "OptionsHeader";
            this.OptionsHeader.Size = new System.Drawing.Size(112, 19);
            this.OptionsHeader.TabIndex = 8;
            this.OptionsHeader.Text = "Sync Options";
            // 
            // TwoWaySyncDescription1
            // 
            this.TwoWaySyncDescription1.AutoSize = true;
            this.TwoWaySyncDescription1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TwoWaySyncDescription1.Location = new System.Drawing.Point(106, 268);
            this.TwoWaySyncDescription1.Name = "TwoWaySyncDescription1";
            this.TwoWaySyncDescription1.Size = new System.Drawing.Size(327, 14);
            this.TwoWaySyncDescription1.TabIndex = 10;
            this.TwoWaySyncDescription1.Text = "Sync both your Google Contacts and Microsoft Outlook cards with";
            // 
            // TwoWaySyncDescription2
            // 
            this.TwoWaySyncDescription2.AutoSize = true;
            this.TwoWaySyncDescription2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TwoWaySyncDescription2.Location = new System.Drawing.Point(106, 282);
            this.TwoWaySyncDescription2.Name = "TwoWaySyncDescription2";
            this.TwoWaySyncDescription2.Size = new System.Drawing.Size(59, 14);
            this.TwoWaySyncDescription2.TabIndex = 11;
            this.TwoWaySyncDescription2.Text = "each other";
            // 
            // GoogleToOutlook
            // 
            this.GoogleToOutlook.AutoSize = true;
            this.GoogleToOutlook.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GoogleToOutlook.Location = new System.Drawing.Point(90, 305);
            this.GoogleToOutlook.Name = "GoogleToOutlook";
            this.GoogleToOutlook.Size = new System.Drawing.Size(361, 20);
            this.GoogleToOutlook.TabIndex = 12;
            this.GoogleToOutlook.Text = "1-way: Google Contacts to Microsoft Outlook contacts";
            this.GoogleToOutlook.UseVisualStyleBackColor = true;
            // 
            // GoogleToOutlookDescription
            // 
            this.GoogleToOutlookDescription.AutoSize = true;
            this.GoogleToOutlookDescription.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GoogleToOutlookDescription.Location = new System.Drawing.Point(106, 324);
            this.GoogleToOutlookDescription.Name = "GoogleToOutlookDescription";
            this.GoogleToOutlookDescription.Size = new System.Drawing.Size(350, 14);
            this.GoogleToOutlookDescription.TabIndex = 13;
            this.GoogleToOutlookDescription.Text = "Sync only your Google Contacts cards with Microsoft Outlook contacts";
            // 
            // OutlookToGoogle
            // 
            this.OutlookToGoogle.AutoSize = true;
            this.OutlookToGoogle.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OutlookToGoogle.Location = new System.Drawing.Point(90, 347);
            this.OutlookToGoogle.Name = "OutlookToGoogle";
            this.OutlookToGoogle.Size = new System.Drawing.Size(365, 20);
            this.OutlookToGoogle.TabIndex = 14;
            this.OutlookToGoogle.Text = "1-way: Microsoft Outlook contacts to Google Contacts ";
            this.OutlookToGoogle.UseVisualStyleBackColor = true;
            // 
            // OutlookToGoogleDescription
            // 
            this.OutlookToGoogleDescription.AutoSize = true;
            this.OutlookToGoogleDescription.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OutlookToGoogleDescription.Location = new System.Drawing.Point(106, 366);
            this.OutlookToGoogleDescription.Name = "OutlookToGoogleDescription";
            this.OutlookToGoogleDescription.Size = new System.Drawing.Size(305, 14);
            this.OutlookToGoogleDescription.TabIndex = 15;
            this.OutlookToGoogleDescription.Text = "Sync only your Microsoft Outlook cards with Google Contacts";
            // 
            // Save
            // 
            this.Save.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.Save.Location = new System.Drawing.Point(320, 454);
            this.Save.Name = "Save";
            this.Save.Size = new System.Drawing.Size(70, 24);
            this.Save.TabIndex = 21;
            this.Save.Text = "Save";
            this.Save.UseVisualStyleBackColor = true;
            this.Save.Click += new System.EventHandler(this.Save_Click);
            // 
            // GoogleContactsSignup
            // 
            this.GoogleContactsSignup.AutoSize = true;
            this.GoogleContactsSignup.Location = new System.Drawing.Point(106, 196);
            this.GoogleContactsSignup.Name = "GoogleContactsSignup";
            this.GoogleContactsSignup.Size = new System.Drawing.Size(171, 16);
            this.GoogleContactsSignup.TabIndex = 7;
            this.GoogleContactsSignup.TabStop = true;
            this.GoogleContactsSignup.Text = "Sign up for Google Contacts";
            this.GoogleContactsSignup.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.GoogleContactsSignup_LinkClicked);
            // 
            // Cancel
            // 
            this.Cancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel.Location = new System.Drawing.Point(399, 454);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(70, 24);
            this.Cancel.TabIndex = 22;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // Help
            // 
            this.Help.AutoSize = true;
            this.Help.Location = new System.Drawing.Point(211, 460);
            this.Help.Name = "Help";
            this.Help.Size = new System.Drawing.Size(34, 16);
            this.Help.TabIndex = 20;
            this.Help.TabStop = true;
            this.Help.Text = "Help";
            this.Help.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.Help_LinkClicked);
            // 
            // UserName
            // 
            this.UserName.Location = new System.Drawing.Point(109, 134);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(224, 22);
            this.UserName.TabIndex = 4;
            // 
            // Worker
            // 
            this.Worker.WorkerReportsProgress = true;
            this.Worker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.Worker_DoWork);
            this.Worker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.Worker_ProgressChanged);
            this.Worker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.Worker_RunWorkerCompleted);
            // 
            // WorkTimer
            // 
            this.WorkTimer.Tick += new System.EventHandler(this.WorkTimer_Tick);
            // 
            // DoubleClickTimer
            // 
            this.DoubleClickTimer.Tick += new System.EventHandler(this.DoubleClickTimer_Tick);
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.Save;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.CancelButton = this.Cancel;
            this.ClientSize = new System.Drawing.Size(488, 494);
            this.Controls.Add(this.Help);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.GoogleContactsSignup);
            this.Controls.Add(this.Save);
            this.Controls.Add(this.OutlookToGoogleDescription);
            this.Controls.Add(this.OutlookToGoogle);
            this.Controls.Add(this.GoogleToOutlookDescription);
            this.Controls.Add(this.GoogleToOutlook);
            this.Controls.Add(this.TwoWaySyncDescription2);
            this.Controls.Add(this.TwoWaySyncDescription1);
            this.Controls.Add(this.OptionsHeader);
            this.Controls.Add(this.Description2);
            this.Controls.Add(this.Description1);
            this.Controls.Add(this.GoogleContactsImage);
            this.Controls.Add(this.SyncInterval);
            this.Controls.Add(this.SyncIntervalUnit);
            this.Controls.Add(this.TwoWaySync);
            this.Controls.Add(this.SyncIntervalLabel);
            this.Controls.Add(this.UserName);
            this.Controls.Add(this.AccountHeader);
            this.Controls.Add(this.PasswordLabel);
            this.Controls.Add(this.UserNameLabel);
            this.Controls.Add(this.Password);
            this.Controls.Add(this.ResetMatches);
            this.Controls.Add(this.ContactsImage);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.Text = "Google Contacts Sync {0}";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.Shown += new System.EventHandler(this.SettingsForm_Shown);
            this.NotificationMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ContactsImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.SyncInterval)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GoogleContactsImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label PasswordLabel;
        private System.Windows.Forms.Label UserNameLabel;
        private System.Windows.Forms.Timer SyncTimer;
        private System.Windows.Forms.ContextMenuStrip NotificationMenu;
        private System.Windows.Forms.ToolStripMenuItem OptionsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExitMenuItem;
        private System.Windows.Forms.ToolStripSeparator SecondMenuSeparator;
        private System.Windows.Forms.ToolStripMenuItem SyncMenuItem;
        private System.Windows.Forms.ToolStripSeparator FirstMenuSeparator;
        private System.Windows.Forms.LinkLabel ResetMatches;
        private System.Windows.Forms.PictureBox ContactsImage;
        private System.Windows.Forms.NumericUpDown SyncInterval;
        private System.Windows.Forms.Label SyncIntervalLabel;
        private System.Windows.Forms.RadioButton TwoWaySync;
        private System.Windows.Forms.Label AccountHeader;
        private System.Windows.Forms.PictureBox GoogleContactsImage;
        private System.Windows.Forms.Label Description1;
        private System.Windows.Forms.Label Description2;
        private System.Windows.Forms.Label OptionsHeader;
        private System.Windows.Forms.Label TwoWaySyncDescription1;
        private System.Windows.Forms.Label SyncIntervalUnit;
        private System.Windows.Forms.Label TwoWaySyncDescription2;
        private System.Windows.Forms.RadioButton GoogleToOutlook;
        private System.Windows.Forms.Label GoogleToOutlookDescription;
        private System.Windows.Forms.RadioButton OutlookToGoogle;
        private System.Windows.Forms.Label OutlookToGoogleDescription;
        private System.Windows.Forms.Button Save;
        private System.Windows.Forms.LinkLabel GoogleContactsSignup;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.LinkLabel Help;
        private System.Windows.Forms.NotifyIcon Notifications;
        private System.ComponentModel.BackgroundWorker Worker;
        private System.Windows.Forms.Timer WorkTimer;
        private System.Windows.Forms.Timer DoubleClickTimer;

    }
}

