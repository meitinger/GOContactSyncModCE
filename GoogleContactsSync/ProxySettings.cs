using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Net;

namespace GoContactSyncMod
{
    partial class ProxySettingsForm : Form
    {

        private void Form_Changed(object sender, EventArgs e)
        {
            FormSettings();
        }

        private void setBgColor(TextBox box, bool isValid)
        {
            if (box.Enabled)
            {
                if (!isValid)
                    box.BackColor = Color.LightPink;
                else
                    box.BackColor = Color.LightGreen;
            }
        }

        private bool ValidCredentials
        {
            get
            {
                bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\@\'\%\._\+\-]+)$", RegexOptions.IgnoreCase);
                bool passwordIsValid = Password.Text.Length != 0;
                bool AddressIsValid  = Regex.IsMatch(Address.Text, @"^(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
                bool PortIsValid     = Regex.IsMatch(Port.Text, @"^(?'port'[0-9]{2,6})$", RegexOptions.IgnoreCase);


                setBgColor(UserName, userNameIsValid);
                setBgColor(Password, passwordIsValid);
                setBgColor(Address,  AddressIsValid);
                setBgColor(Port,     PortIsValid);
                return (userNameIsValid && passwordIsValid || !Authorization.Checked) && AddressIsValid && PortIsValid || SystemProxy.Checked;
            }
        }

        private void FormSettings()
        {
            Address.Enabled       = CustomProxy.Checked;
            Port.Enabled          = CustomProxy.Checked;
            Authorization.Enabled = CustomProxy.Checked; 
            UserName.Enabled      = CustomProxy.Checked && Authorization.Checked;
            Password.Enabled      = CustomProxy.Checked && Authorization.Checked;

            bool isValid = ValidCredentials;
        }

        private void ProxySet()
        {
            if (CustomProxy.Checked)
            {
                System.Net.WebProxy myProxy = new System.Net.WebProxy(Address.Text, Convert.ToInt16(Port.Text));
                myProxy.BypassProxyOnLocal = false;

                if (Authorization.Checked)
                {
                    myProxy.Credentials = new System.Net.NetworkCredential(UserName.Text, Password.Text);
                }
                GlobalProxySelection.Select = myProxy;
            }
            else
                GlobalProxySelection.Select = null;
        }

        public ProxySettingsForm()
        {
            InitializeComponent();
            LoadSettings();
            
            ProxySet();
        }

        private void LoadSettings()
        {   // Load Proxy Settings
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            if (regKeyAppRoot.GetValue("ProxyUsage") != null)
            {
                if (Convert.ToBoolean (regKeyAppRoot.GetValue("ProxyUsage")))
                {
                    CustomProxy.Checked = true;
                    SystemProxy.Checked = !CustomProxy.Checked;

                    if (regKeyAppRoot.GetValue("ProxyURL") != null)
                        Address.Text = (string)regKeyAppRoot.GetValue("ProxyURL");

                    if (regKeyAppRoot.GetValue("ProxyPort") != null)
                        Port.Text = (string)regKeyAppRoot.GetValue("ProxyPort");

                    if (Convert.ToBoolean (regKeyAppRoot.GetValue("ProxyAuth"))) 
                    {

                         Authorization.Checked = true;

                        if (regKeyAppRoot.GetValue("ProxyUsername") != null)
                        {
                            UserName.Text = regKeyAppRoot.GetValue("ProxyUsername") as string;
                            if (regKeyAppRoot.GetValue("ProxyPassword") != null)
                                Password.Text = Encryption.DecryptPassword(UserName.Text, regKeyAppRoot.GetValue("ProxyPassword") as string);
                        }
                    }
                }
            }

            FormSettings();
        }

        private void SaveSettings()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            regKeyAppRoot.SetValue("ProxyUsage", CustomProxy.Checked);

            if (CustomProxy.Checked)
            {

                if (!string.IsNullOrEmpty(Address.Text))
                {
                    regKeyAppRoot.SetValue("ProxyURL", Address.Text);
                    if (!string.IsNullOrEmpty(Port.Text))
                        regKeyAppRoot.SetValue("ProxyPort", Port.Text);
                }

                regKeyAppRoot.SetValue("ProxyAuth", Authorization.Checked);
                if (Authorization.Checked) 
                {
                    if (!string.IsNullOrEmpty(UserName.Text))
                    {
                        regKeyAppRoot.SetValue("ProxyUsername", UserName.Text);
                        if (!string.IsNullOrEmpty(Password.Text))
                            regKeyAppRoot.SetValue("ProxyPassword", Encryption.EncryptPassword(UserName.Text, Password.Text));
                    }
                }
            }
        }


        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (!ValidCredentials)
                return;

            SaveSettings();
            ProxySet();
            this.Close();
        }

    }
}

