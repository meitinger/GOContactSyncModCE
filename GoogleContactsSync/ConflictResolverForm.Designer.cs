namespace GoContactSyncMod
{
    partial class ConflictResolverForm
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
            this.messageLabel = new System.Windows.Forms.Label();
            this.keepOutlook = new System.Windows.Forms.Button();
            this.keepGoogle = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.skip = new System.Windows.Forms.Button();
            this.OutlookItemTextBox = new System.Windows.Forms.TextBox();
            this.SplitContainer = new System.Windows.Forms.SplitContainer();
            this.OutlookItemLabel = new System.Windows.Forms.Label();
            this.GoogleItemTextBox = new System.Windows.Forms.TextBox();
            this.GoogleItemLabel = new System.Windows.Forms.Label();
            this.AllCheckBox = new System.Windows.Forms.CheckBox();
            this.GoogleComboBox = new System.Windows.Forms.ComboBox();
            this.SplitContainer.Panel1.SuspendLayout();
            this.SplitContainer.Panel2.SuspendLayout();
            this.SplitContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // messageLabel
            // 
            this.messageLabel.Location = new System.Drawing.Point(12, 9);
            this.messageLabel.Name = "messageLabel";
            this.messageLabel.Size = new System.Drawing.Size(417, 62);
            this.messageLabel.TabIndex = 0;
            this.messageLabel.Text = "message";
            // 
            // keepOutlook
            // 
            this.keepOutlook.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.keepOutlook.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.keepOutlook.Location = new System.Drawing.Point(12, 320);
            this.keepOutlook.Name = "keepOutlook";
            this.keepOutlook.Size = new System.Drawing.Size(120, 23);
            this.keepOutlook.TabIndex = 1;
            this.keepOutlook.Text = "Keep &Outlook Entry";
            this.keepOutlook.UseVisualStyleBackColor = true;
            // 
            // keepGoogle
            // 
            this.keepGoogle.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.keepGoogle.DialogResult = System.Windows.Forms.DialogResult.No;
            this.keepGoogle.Location = new System.Drawing.Point(138, 320);
            this.keepGoogle.Name = "keepGoogle";
            this.keepGoogle.Size = new System.Drawing.Size(125, 23);
            this.keepGoogle.TabIndex = 2;
            this.keepGoogle.Text = "Keep &Google Entry";
            this.keepGoogle.UseVisualStyleBackColor = true;
            // 
            // cancel
            // 
            this.cancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancel.Location = new System.Drawing.Point(350, 320);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(75, 23);
            this.cancel.TabIndex = 3;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = true;
            // 
            // skip
            // 
            this.skip.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.skip.DialogResult = System.Windows.Forms.DialogResult.Ignore;
            this.skip.Location = new System.Drawing.Point(269, 320);
            this.skip.Name = "skip";
            this.skip.Size = new System.Drawing.Size(75, 23);
            this.skip.TabIndex = 3;
            this.skip.Text = "&Skip";
            this.skip.UseVisualStyleBackColor = true;
            // 
            // OutlookItemTextBox
            // 
            this.OutlookItemTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.OutlookItemTextBox.Location = new System.Drawing.Point(0, 13);
            this.OutlookItemTextBox.Multiline = true;
            this.OutlookItemTextBox.Name = "OutlookItemTextBox";
            this.OutlookItemTextBox.ReadOnly = true;
            this.OutlookItemTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.OutlookItemTextBox.Size = new System.Drawing.Size(206, 204);
            this.OutlookItemTextBox.TabIndex = 4;
            // 
            // SplitContainer
            // 
            this.SplitContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.SplitContainer.Location = new System.Drawing.Point(15, 74);
            this.SplitContainer.Name = "SplitContainer";
            // 
            // SplitContainer.Panel1
            // 
            this.SplitContainer.Panel1.Controls.Add(this.OutlookItemTextBox);
            this.SplitContainer.Panel1.Controls.Add(this.OutlookItemLabel);
            // 
            // SplitContainer.Panel2
            // 
            this.SplitContainer.Panel2.Controls.Add(this.GoogleItemTextBox);
            this.SplitContainer.Panel2.Controls.Add(this.GoogleItemLabel);
            this.SplitContainer.Size = new System.Drawing.Size(410, 217);
            this.SplitContainer.SplitterDistance = 206;
            this.SplitContainer.TabIndex = 5;
            // 
            // OutlookItemLabel
            // 
            this.OutlookItemLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.OutlookItemLabel.Location = new System.Drawing.Point(0, 0);
            this.OutlookItemLabel.Name = "OutlookItemLabel";
            this.OutlookItemLabel.Size = new System.Drawing.Size(206, 13);
            this.OutlookItemLabel.TabIndex = 6;
            this.OutlookItemLabel.Text = "Outlook";
            // 
            // GoogleItemTextBox
            // 
            this.GoogleItemTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.GoogleItemTextBox.Location = new System.Drawing.Point(0, 13);
            this.GoogleItemTextBox.Multiline = true;
            this.GoogleItemTextBox.Name = "GoogleItemTextBox";
            this.GoogleItemTextBox.ReadOnly = true;
            this.GoogleItemTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.GoogleItemTextBox.Size = new System.Drawing.Size(200, 204);
            this.GoogleItemTextBox.TabIndex = 5;
            // 
            // GoogleItemLabel
            // 
            this.GoogleItemLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.GoogleItemLabel.Location = new System.Drawing.Point(0, 0);
            this.GoogleItemLabel.Name = "GoogleItemLabel";
            this.GoogleItemLabel.Size = new System.Drawing.Size(200, 13);
            this.GoogleItemLabel.TabIndex = 7;
            this.GoogleItemLabel.Text = "Google";
            // 
            // AllCheckBox
            // 
            this.AllCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.AllCheckBox.AutoSize = true;
            this.AllCheckBox.Location = new System.Drawing.Point(15, 297);
            this.AllCheckBox.Name = "AllCheckBox";
            this.AllCheckBox.Size = new System.Drawing.Size(150, 17);
            this.AllCheckBox.TabIndex = 6;
            this.AllCheckBox.Text = "For all following items";
            this.AllCheckBox.UseVisualStyleBackColor = true;
            // 
            // GoogleComboBox
            // 
            this.GoogleComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.GoogleComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.GoogleComboBox.FormattingEnabled = true;
            this.GoogleComboBox.Location = new System.Drawing.Point(225, 293);
            this.GoogleComboBox.Name = "GoogleComboBox";
            this.GoogleComboBox.Size = new System.Drawing.Size(200, 21);
            this.GoogleComboBox.Sorted = true;
            this.GoogleComboBox.TabIndex = 7;
            this.GoogleComboBox.Visible = false;
            this.GoogleComboBox.SelectedIndexChanged += new System.EventHandler(this.GoogleComboBox_SelectedIndexChanged);
            // 
            // ConflictResolverForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancel;
            this.ClientSize = new System.Drawing.Size(434, 355);
            this.Controls.Add(this.GoogleComboBox);
            this.Controls.Add(this.AllCheckBox);
            this.Controls.Add(this.SplitContainer);
            this.Controls.Add(this.skip);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.keepGoogle);
            this.Controls.Add(this.keepOutlook);
            this.Controls.Add(this.messageLabel);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConflictResolverForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Confict detected!";
            this.SplitContainer.Panel1.ResumeLayout(false);
            this.SplitContainer.Panel1.PerformLayout();
            this.SplitContainer.Panel2.ResumeLayout(false);
            this.SplitContainer.Panel2.PerformLayout();
            this.SplitContainer.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Button keepOutlook;
        public System.Windows.Forms.Button keepGoogle;
        private System.Windows.Forms.Button cancel;
        public System.Windows.Forms.Label messageLabel;
        public System.Windows.Forms.Button skip;
        public System.Windows.Forms.TextBox OutlookItemTextBox;
        private System.Windows.Forms.SplitContainer SplitContainer;
        private System.Windows.Forms.Label OutlookItemLabel;
        public System.Windows.Forms.TextBox GoogleItemTextBox;
        private System.Windows.Forms.Label GoogleItemLabel;
        public System.Windows.Forms.CheckBox AllCheckBox;
        public System.Windows.Forms.ComboBox GoogleComboBox;
    }
}