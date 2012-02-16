using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class ConfigurationManagerForm : Form
    {
        public ConfigurationManagerForm()
        {
            InitializeComponent();
        }

        public string AddProfile()
        {
            string vReturn = "";
            AddEditProfileForm AddEditProfile = new AddEditProfileForm("New profile", null);
            if (AddEditProfile.ShowDialog() == DialogResult.OK)
            {
                if (null != Registry.CurrentUser.OpenSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName))
                {
                    MessageBox.Show("Profile " + AddEditProfile.ProfileName + " exists, try again. ", "New profile");
                }
                else
                {
                    Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName);
                    vReturn = AddEditProfile.ProfileName;
                }
            }

            return vReturn;
        }

        private void fillListProfiles()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey);

            lbProfiles.Items.Clear();

            foreach (string subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                lbProfiles.Items.Add(subKeyName);
            }
        }

        //copy all the values
        private void CopyKey(RegistryKey parent, string keyNameSource, string keyNameDestination)
        {
            RegistryKey destination = parent.CreateSubKey(keyNameDestination);
            RegistryKey source = parent.OpenSubKey(keyNameSource);

            foreach (string valueName in source.GetValueNames())
            {
                object objValue = source.GetValue(valueName);
                RegistryValueKind valKind = source.GetValueKind(valueName);
                destination.SetValue(valueName, objValue, valKind);
            }
        }

        private void btClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btAdd_Click(object sender, EventArgs e)
        {
            AddProfile();
            fillListProfiles();
        }

        private void btEdit_Click(object sender, EventArgs e)
        {
            if (1 == lbProfiles.CheckedItems.Count)
            {
                AddEditProfileForm AddEditProfile = new AddEditProfileForm("Edit profile",lbProfiles.CheckedItems[0].ToString());
                if (AddEditProfile.ShowDialog() == DialogResult.OK)
                {
                    if (null != Registry.CurrentUser.OpenSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName))
                    {
                        MessageBox.Show("Profile " + AddEditProfile.ProfileName + " exists, try again. ", "Edit profile");
                    }
                    else 
                    {
                        CopyKey(Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey), lbProfiles.CheckedItems[0].ToString(), AddEditProfile.ProfileName);
                        Registry.CurrentUser.DeleteSubKeyTree(SettingsForm.AppRootKey + '\\' + lbProfiles.CheckedItems[0].ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, select one profile for editing", "Edit profile");
            }

            fillListProfiles();
        }

        private void btDel_Click(object sender, EventArgs e)
        {
            if (0 >= lbProfiles.CheckedItems.Count) 
            {  
                 MessageBox.Show("You don`t select any profile. Deletion imposble.", "Delete profile");
            } else if (DialogResult.Yes == MessageBox.Show("Do you sure to delete selection ?", "Delete profile",
                                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                foreach (object itemChecked in lbProfiles.CheckedItems)
                {
                    Registry.CurrentUser.DeleteSubKeyTree(SettingsForm.AppRootKey + '\\' + itemChecked.ToString());
                }
            }

            fillListProfiles();
        }

        private void ConfigurationManagerForm_Load(object sender, EventArgs e)
        {
            fillListProfiles();
        }
    }
}
