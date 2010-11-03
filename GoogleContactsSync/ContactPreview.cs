using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;

namespace WebGear.GoogleContactsSync
{
    public partial class ContactPreview : UserControl
    {
        private Collection<CPField> fields;

        private Outlook.ContactItem outlookContact;
        public Outlook.ContactItem OutlookContact
        {
            get { return outlookContact; }
            set { outlookContact = value; }
        }

        public ContactPreview(Outlook.ContactItem _outlookContact)
        {
            InitializeComponent();
            outlookContact = _outlookContact;
            InitializeFields();
        }

        private void InitializeFields()
        {
            // TODO: init all non null fields
            fields = new Collection<CPField>();

            int index = 0;
            int height = Font.Height;

            if (outlookContact.FirstName != null)
            {
                fields.Add(new CPField("First name", outlookContact.FirstName, new PointF(0, index * height)));
                index++;
            }
            if (outlookContact.LastName != null)
            {
                fields.Add(new CPField("Last name", outlookContact.LastName, new PointF(0, index * height)));
                index++;
            }
            if (outlookContact.Email1Address != null)
            {
                fields.Add(new CPField("Email", outlookContact.Email1Address, new PointF(0, index * height)));
                index++;
            }
            
            // resize to fit
            this.Height = (index + 1) * height;
        }

        private void ContactPreview_Paint(object sender, PaintEventArgs e)
        {
            foreach (CPField field in fields)
                field.Draw(e, Font);
        }


    }

    public class CPField
    {
        public string name;
        public string value;
        public PointF p;

        public CPField(string _name, string _value, PointF _p)
        {
            name = _name;
            value = _value;
            p = _p;
        }

        public void Draw(PaintEventArgs e, Font font)
        {
            string str = name + ": " + value;
            e.Graphics.DrawString(str, font, Brushes.Black, p);
        }
    }
}
