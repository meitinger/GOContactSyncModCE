using System;
using System.Collections.Generic;
using System.Text;

namespace WebGear.GoogleContactsSync
{
    class ConflictResolver : IConflictResolver
    {
        private ConflictResolverForm _form;

        public ConflictResolver()
        {
            _form = new ConflictResolverForm();
        }
        public ConflictResolver(ConflictResolverForm form)
        {
            _form = form;
        }



        #region IConflictResolver Members

        public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.ContactItem outlookContact, Google.GData.Contacts.ContactEntry googleContact)
        {
            _form.messageLabel.Text =
                "Both the outlook contact and the google contact \"" + googleContact.Title.Text +
                "\" has been changed. Choose which you would like to keep.";

            switch (_form.ShowDialog())
            {
                case System.Windows.Forms.DialogResult.Ignore:
                    // skip
                    return ConflictResolution.Skip;
                case System.Windows.Forms.DialogResult.Cancel:
                    // cancel
                    return ConflictResolution.Cancel;
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return ConflictResolution.GoogleWins;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return ConflictResolution.OutlookWins;
                default:
                    throw new Exception();
            }
        }

        #endregion
    }
}
