using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;

namespace WebGear.GoogleContactsSync
{
    public partial class DeleteDuplicatesForm : Form
    {
        private Collection<ContactPreview> previews;

        public DeleteDuplicatesForm(Collection<Outlook.ContactItem> outlookContacts)
        {
            InitializeComponent();
            previews = new Collection<ContactPreview>();
            Collection<Outlook.ContactItem> duplicates = FindDuplicates(outlookContacts);

            foreach (Outlook.ContactItem outlookContact in duplicates)
            {
                ContactPreview preview = new ContactPreview(outlookContact);
                preview.Parent = flowLayoutPanel;
                previews.Add(preview);
            }
        }

        private static Collection<Outlook.ContactItem> FindDuplicates(Collection<Outlook.ContactItem> outlookContacts)
        {
            // TODO:
            return null;
        }
    }
}