using System;
using System.Collections.Generic;
using System.Text;
using Google.Contacts;
using Google.Documents;
using System.Reflection;
using Google.GData.Extensions;

namespace GoContactSyncMod
{
    class ConflictResolver : IConflictResolver
    {
        private ConflictResolverForm _form;

        public ConflictResolver()
        {
            _form = new ConflictResolverForm();
        }
        

        #region IConflictResolver Members

        public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.ContactItem outlookContact, Contact googleContact)
        {
            string name = String.Empty;
            if (googleContact != null)
                ContactMatch.GetGoogleContactName(googleContact);
           
            _form.messageLabel.Text =
                "Both the outlook contact and the google contact \"" + name +
                "\" have been changed. Choose which you would like to keep.";
            
            _form.OutlookItemTextBox.Text = "Name: " + outlookContact.FileAs + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Email1Address))
                _form.OutlookItemTextBox.Text += "Email1: " + outlookContact.Email1Address + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Email2Address))
                _form.OutlookItemTextBox.Text += "Email2: " + outlookContact.Email2Address + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Email3Address))
                _form.OutlookItemTextBox.Text += "Email3: " + outlookContact.Email3Address + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.MobileTelephoneNumber))
                _form.OutlookItemTextBox.Text += "MobilePhone: " + outlookContact.MobileTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.HomeTelephoneNumber))
                _form.OutlookItemTextBox.Text += "HomePhone: " + outlookContact.HomeTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Home2TelephoneNumber))
                _form.OutlookItemTextBox.Text += "HomePhone2: " + outlookContact.HomeTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.BusinessTelephoneNumber))
               _form.OutlookItemTextBox.Text += "BusinessPhone: " + outlookContact.BusinessTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.Business2TelephoneNumber))
               _form.OutlookItemTextBox.Text += "BusinessPhone2: " + outlookContact.BusinessTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.OtherTelephoneNumber))                
                _form.OutlookItemTextBox.Text += "OtherPhone: " + outlookContact.OtherTelephoneNumber + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.HomeAddress))
                _form.OutlookItemTextBox.Text += "HomeAddress: " + outlookContact.HomeAddress + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.BusinessAddress))
                _form.OutlookItemTextBox.Text += "BusinessAddress: " + outlookContact.BusinessAddress + "\r\n";
            if (!string.IsNullOrEmpty(outlookContact.OtherAddress))
                _form.OutlookItemTextBox.Text += "OtherAddress: " + outlookContact.OtherAddress + "\r\n";

            _form.GoogleItemTextBox.Text = "Name: " + googleContact.Name.FullName + "\r\n";
            for (int i = 0; i < googleContact.Emails.Count; i++)
            {
                string email = googleContact.Emails[i].Address;
                if (!string.IsNullOrEmpty(email))
                {
                    _form.GoogleItemTextBox.Text += "Email" + (i+1) + ": " + email + "\r\n";
                }
            }
            foreach (PhoneNumber phone in googleContact.Phonenumbers)
            {
                if (!string.IsNullOrEmpty(phone.Value))
                {
                    if (phone.Rel == ContactsRelationships.IsMobile)
                        _form.GoogleItemTextBox.Text += "MobilePhone: ";
                    if (phone.Rel == ContactsRelationships.IsHome)
                        _form.GoogleItemTextBox.Text += "HomePhone: ";
                    if (phone.Rel == ContactsRelationships.IsWork)
                        _form.GoogleItemTextBox.Text += "BusinessPhone: ";
                    if (phone.Rel == ContactsRelationships.IsOther)
                        _form.GoogleItemTextBox.Text += "OtherPhone: ";

                    _form.GoogleItemTextBox.Text += phone.Value + "\r\n";
                }

                
            }

            foreach (StructuredPostalAddress address in googleContact.PostalAddresses)
            {
                if (!string.IsNullOrEmpty(address.FormattedAddress))
                {
                    if (address.Rel == ContactsRelationships.IsHome)
                        _form.GoogleItemTextBox.Text += "HomeAddress: ";
                    if (address.Rel == ContactsRelationships.IsWork)
                        _form.GoogleItemTextBox.Text += "BusinessAddress: ";
                    if (address.Rel == ContactsRelationships.IsOther)
                        _form.GoogleItemTextBox.Text += "OtherAddress: ";   

                     _form.GoogleItemTextBox.Text += address.FormattedAddress + "\r\n";
                }
            }

            return Resolve();
        }

        private ConflictResolution Resolve()
        {

            switch (_form.ShowDialog())
            {
                case System.Windows.Forms.DialogResult.Ignore:
                    // skip
                    return _form.AllCheckBox.Checked ? ConflictResolution.SkipAlways : ConflictResolution.Skip;
                case System.Windows.Forms.DialogResult.Cancel:
                    // cancel
                    return ConflictResolution.Cancel;
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.GoogleWinsAlways : ConflictResolution.GoogleWins;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.OutlookWinsAlways : ConflictResolution.OutlookWins;
                default:
                    throw new Exception();
            }
        }

        //private string GetPropertyInfos(object myObject)
        //{
        //    // get all public static properties of OutlookItem
        //    PropertyInfo[] propertyInfos;
        //    propertyInfos = myObject.GetType().GetProperties();
        //    // sort properties by name
        //    Array.Sort(propertyInfos,
        //            delegate(PropertyInfo propertyInfo1, PropertyInfo propertyInfo2)
        //            { return propertyInfo1.Name.CompareTo(propertyInfo2.Name); });

        //    string ret = String.Empty;
        //    // write property names
        //    foreach (PropertyInfo propertyInfo in propertyInfos)
        //    {
        //        object value = propertyInfo.GetValue(myObject, null);

        //        if (value != null && !(value is string))
        //            ret += propertyInfo.Name + ": " + value + "\r\n";
        //    }

        //    return ret;
        //}

        public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.NoteItem outlookNote, Document googleNote, Syncronizer sync)
        {
            string name = googleNote.Title;
            
            _form.messageLabel.Text =
                "Both the outlook note and the google note \"" + name +
                "\" have been changed. Choose which you would like to keep.";

            _form.OutlookItemTextBox.Text = outlookNote.Body;
            _form.GoogleItemTextBox.Text = NotePropertiesUtils.GetBody(sync, googleNote);

            return Resolve();
        }

        #endregion
    }
}
