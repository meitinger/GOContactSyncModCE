namespace WebGear.GoogleContactsSync
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
            this.UserName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.syncButton = new System.Windows.Forms.Button();
            this.syncOptionBox = new System.Windows.Forms.CheckedListBox();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.systemTrayMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.autoSyncInterval = new System.Windows.Forms.NumericUpDown();
            this.autoSyncCheckBox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.runAtStartupCheckBox = new System.Windows.Forms.CheckBox();
            this.nextSyncLabel = new System.Windows.Forms.Label();
            this.syncTimer = new System.Windows.Forms.Timer(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btSyncDelete = new System.Windows.Forms.CheckBox();
            this.tbSyncProfile = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lastSyncLabel = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.syncConsole = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.resetMatchesButton = new System.Windows.Forms.Button();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.systemTrayMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // Password
            // 
            this.Password.Location = new System.Drawing.Point(86, 44);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.Size = new System.Drawing.Size(182, 20);
            this.Password.TabIndex = 16;
            this.Password.TextChanged += new System.EventHandler(this.Password_TextChanged);
            // 
            // UserName
            // 
            this.UserName.Location = new System.Drawing.Point(86, 18);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(182, 20);
            this.UserName.TabIndex = 15;
            this.UserName.TextChanged += new System.EventHandler(this.UserName_TextChanged);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(6, 44);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 16);
            this.label3.TabIndex = 14;
            this.label3.Text = "Password:";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(6, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 17);
            this.label2.TabIndex = 13;
            this.label2.Text = "User:";
            // 
            // syncButton
            // 
            this.syncButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.syncButton.Location = new System.Drawing.Point(210, 510);
            this.syncButton.Name = "syncButton";
            this.syncButton.Size = new System.Drawing.Size(75, 23);
            this.syncButton.TabIndex = 19;
            this.syncButton.Text = "Sync";
            this.syncButton.UseVisualStyleBackColor = true;
            this.syncButton.Click += new System.EventHandler(this.button4_Click);
            // 
            // syncOptionBox
            // 
            this.syncOptionBox.BackColor = System.Drawing.SystemColors.Control;
            this.syncOptionBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.syncOptionBox.CheckOnClick = true;
            this.syncOptionBox.FormattingEnabled = true;
            this.syncOptionBox.Location = new System.Drawing.Point(6, 73);
            this.syncOptionBox.Name = "syncOptionBox";
            this.syncOptionBox.Size = new System.Drawing.Size(262, 75);
            this.syncOptionBox.TabIndex = 25;
            this.toolTip.SetToolTip(this.syncOptionBox, resources.GetString("syncOptionBox.ToolTip"));
            this.syncOptionBox.SelectedIndexChanged += new System.EventHandler(this.syncOptionBox_SelectedIndexChanged);
            // 
            // notifyIcon
            // 
            this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning;
            this.notifyIcon.ContextMenuStrip = this.systemTrayMenu;
            this.notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
            this.notifyIcon.Text = "GO Contact Sync";
            this.notifyIcon.Visible = true;
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseDoubleClick);
            // 
            // systemTrayMenu
            // 
            this.systemTrayMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem4,
            this.toolStripSeparator2,
            this.toolStripMenuItem1,
            this.toolStripMenuItem3,
            this.toolStripSeparator1,
            this.toolStripMenuItem5,
            this.toolStripMenuItem2});
            this.systemTrayMenu.Name = "systemTrayMenu";
            this.systemTrayMenu.Size = new System.Drawing.Size(116, 126);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(115, 22);
            this.toolStripMenuItem4.Text = "Sync";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.toolStripMenuItem4_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(112, 6);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(115, 22);
            this.toolStripMenuItem1.Text = "Show";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(115, 22);
            this.toolStripMenuItem3.Text = "Hide";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.toolStripMenuItem3_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(112, 6);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(115, 22);
            this.toolStripMenuItem5.Text = "About";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.toolStripMenuItem5_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(115, 22);
            this.toolStripMenuItem2.Text = "Exit";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // autoSyncInterval
            // 
            this.autoSyncInterval.Location = new System.Drawing.Point(93, 64);
            this.autoSyncInterval.Maximum = new decimal(new int[] {
            1440,
            0,
            0,
            0});
            this.autoSyncInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.autoSyncInterval.Name = "autoSyncInterval";
            this.autoSyncInterval.Size = new System.Drawing.Size(42, 20);
            this.autoSyncInterval.TabIndex = 26;
            this.autoSyncInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.autoSyncInterval.Value = new decimal(new int[] {
            120,
            0,
            0,
            0});
            // 
            // autoSyncCheckBox
            // 
            this.autoSyncCheckBox.AutoSize = true;
            this.autoSyncCheckBox.Location = new System.Drawing.Point(12, 42);
            this.autoSyncCheckBox.Name = "autoSyncCheckBox";
            this.autoSyncCheckBox.Size = new System.Drawing.Size(75, 17);
            this.autoSyncCheckBox.TabIndex = 27;
            this.autoSyncCheckBox.Text = "Auto Sync";
            this.autoSyncCheckBox.UseVisualStyleBackColor = true;
            this.autoSyncCheckBox.CheckedChanged += new System.EventHandler(this.autoSyncCheckBox_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 29;
            this.label1.Text = "Sync Interval:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(141, 66);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(28, 13);
            this.label4.TabIndex = 30;
            this.label4.Text = "mins";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.runAtStartupCheckBox);
            this.groupBox1.Controls.Add(this.nextSyncLabel);
            this.groupBox1.Controls.Add(this.autoSyncInterval);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.autoSyncCheckBox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(11, 393);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(274, 111);
            this.groupBox1.TabIndex = 31;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Automization";
            // 
            // runAtStartupCheckBox
            // 
            this.runAtStartupCheckBox.AutoSize = true;
            this.runAtStartupCheckBox.Location = new System.Drawing.Point(12, 21);
            this.runAtStartupCheckBox.Name = "runAtStartupCheckBox";
            this.runAtStartupCheckBox.Size = new System.Drawing.Size(134, 17);
            this.runAtStartupCheckBox.TabIndex = 32;
            this.runAtStartupCheckBox.Text = "Run program at startup";
            this.runAtStartupCheckBox.UseVisualStyleBackColor = true;
            this.runAtStartupCheckBox.CheckedChanged += new System.EventHandler(this.runAtStartupCheckBox_CheckedChanged);
            // 
            // nextSyncLabel
            // 
            this.nextSyncLabel.AutoSize = true;
            this.nextSyncLabel.Location = new System.Drawing.Point(9, 90);
            this.nextSyncLabel.Name = "nextSyncLabel";
            this.nextSyncLabel.Size = new System.Drawing.Size(67, 13);
            this.nextSyncLabel.TabIndex = 31;
            this.nextSyncLabel.Text = "Next Sync in";
            // 
            // syncTimer
            // 
            this.syncTimer.Interval = 1000;
            this.syncTimer.Tick += new System.EventHandler(this.syncTimer_Tick);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.panel1);
            this.groupBox2.Controls.Add(this.btSyncDelete);
            this.groupBox2.Controls.Add(this.tbSyncProfile);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.syncOptionBox);
            this.groupBox2.Location = new System.Drawing.Point(11, 229);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(274, 158);
            this.groupBox2.TabIndex = 32;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Sync Options";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Location = new System.Drawing.Point(6, 67);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(260, 1);
            this.panel1.TabIndex = 36;
            // 
            // btSyncDelete
            // 
            this.btSyncDelete.AutoSize = true;
            this.btSyncDelete.Location = new System.Drawing.Point(7, 46);
            this.btSyncDelete.Name = "btSyncDelete";
            this.btSyncDelete.Size = new System.Drawing.Size(92, 17);
            this.btSyncDelete.TabIndex = 35;
            this.btSyncDelete.Text = "Sync Deletion";
            this.toolTip.SetToolTip(this.btSyncDelete, "This specifies whether deletions are\r\nsynchronized. Enabling this option\r\nmeans i" +
                    "f you delete a contact from\r\nGoogle, then it will be deleted from\r\nOutlook and v" +
                    "ice versa.");
            this.btSyncDelete.UseVisualStyleBackColor = true;
            this.btSyncDelete.CheckedChanged += new System.EventHandler(this.btSyncDelete_CheckedChanged);
            // 
            // tbSyncProfile
            // 
            this.tbSyncProfile.Location = new System.Drawing.Point(86, 19);
            this.tbSyncProfile.Name = "tbSyncProfile";
            this.tbSyncProfile.Size = new System.Drawing.Size(182, 20);
            this.tbSyncProfile.TabIndex = 34;
            this.toolTip.SetToolTip(this.tbSyncProfile, "This is a profile name of your choice.\r\nIt must be unique in each computer\r\nand a" +
                    "ccount you intend to sync with\r\nyour Google Mail account.");
            this.tbSyncProfile.TextChanged += new System.EventHandler(this.tbSyncProfile_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 33;
            this.label5.Text = "Sync Profile:";
            // 
            // lastSyncLabel
            // 
            this.lastSyncLabel.AutoSize = true;
            this.lastSyncLabel.Location = new System.Drawing.Point(6, 16);
            this.lastSyncLabel.Name = "lastSyncLabel";
            this.lastSyncLabel.Size = new System.Drawing.Size(69, 13);
            this.lastSyncLabel.TabIndex = 32;
            this.lastSyncLabel.Text = "Last Sync on";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.syncConsole);
            this.groupBox3.Controls.Add(this.lastSyncLabel);
            this.groupBox3.Location = new System.Drawing.Point(11, 91);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(274, 132);
            this.groupBox3.TabIndex = 33;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Sync Details";
            // 
            // syncConsole
            // 
            this.syncConsole.BackColor = System.Drawing.SystemColors.Info;
            this.syncConsole.Location = new System.Drawing.Point(7, 33);
            this.syncConsole.Multiline = true;
            this.syncConsole.Name = "syncConsole";
            this.syncConsole.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.syncConsole.Size = new System.Drawing.Size(261, 93);
            this.syncConsole.TabIndex = 33;
            this.toolTip.SetToolTip(this.syncConsole, "This window shows information\r\nfrom the last sync.");
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.UserName);
            this.groupBox4.Controls.Add(this.Password);
            this.groupBox4.Location = new System.Drawing.Point(11, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(274, 73);
            this.groupBox4.TabIndex = 34;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Google Account";
            // 
            // resetMatchesButton
            // 
            this.resetMatchesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.resetMatchesButton.AutoSize = true;
            this.resetMatchesButton.Location = new System.Drawing.Point(115, 510);
            this.resetMatchesButton.Name = "resetMatchesButton";
            this.resetMatchesButton.Size = new System.Drawing.Size(89, 23);
            this.resetMatchesButton.TabIndex = 35;
            this.resetMatchesButton.Text = "Reset Matches";
            this.toolTip.SetToolTip(this.resetMatchesButton, "This unlinks Outlook contacts with their\r\ncorresponding Google contatcs. If you\r\n" +
                    "accidentaly delete a contact and you\r\ndont want the deletion to be synchronised," +
                    "\r\nclick  this button.");
            this.resetMatchesButton.UseVisualStyleBackColor = true;
            this.resetMatchesButton.Click += new System.EventHandler(this.resetMatchesButton_Click);
            // 
            // toolTip
            // 
            this.toolTip.IsBalloon = true;
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 537);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.syncButton);
            this.Controls.Add(this.resetMatchesButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.Text = "GO Contact Sync - Settings";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SettingsForm_FormClosed);
            this.Resize += new System.EventHandler(this.SettingsForm_Resize);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.systemTrayMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button syncButton;
        private System.Windows.Forms.CheckedListBox syncOptionBox;
        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.NumericUpDown autoSyncInterval;
        private System.Windows.Forms.CheckBox autoSyncCheckBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Timer syncTimer;
        private System.Windows.Forms.Label nextSyncLabel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lastSyncLabel;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox syncConsole;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button resetMatchesButton;
        private System.Windows.Forms.ContextMenuStrip systemTrayMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.CheckBox runAtStartupCheckBox;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.TextBox tbSyncProfile;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.CheckBox btSyncDelete;
        private System.Windows.Forms.Panel panel1;
    }
}

