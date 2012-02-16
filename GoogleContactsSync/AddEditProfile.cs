using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class AddEditProfileForm : Form
    {
        public string ProfileName
        {
            get { return this.tbProfileName.Text; }
        }
        
        public AddEditProfileForm()
        {
            InitializeComponent();
        }

        public AddEditProfileForm(string title, string profileName)
        {
            InitializeComponent();

            if (!string.IsNullOrEmpty(title))
                this.Text = title;

            if (!string.IsNullOrEmpty(profileName))
                this.tbProfileName.Text = profileName;
        }

    }
}
