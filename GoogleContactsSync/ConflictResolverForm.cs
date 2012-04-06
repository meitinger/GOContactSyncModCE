using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace GoContactSyncMod
{
    internal partial class ConflictResolverForm : Form
    {
        public ConflictResolverForm()
        {
            InitializeComponent();
        }

        private void GoogleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (GoogleComboBox.SelectedItem != null)
                GoogleItemTextBox.Text = ContactMatch.GetSummary((Google.Contacts.Contact)GoogleComboBox.SelectedItem);
        }
    }
}