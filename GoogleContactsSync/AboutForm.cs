using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WebGear.GoogleContactsSync
{
    public partial class AboutForm : Form
    {
        private string urlPath;

        public AboutForm(string _title, string _url,string _text, string _copyright, Image _image)
        {
            InitializeComponent();

            Text += " " + _title;
            title.Text = _title;
            textBox.Text = _text;
            imageBox.Image = _image;
            url.Text = _url;
            urlPath = _url;
            copyrightLabel.Text = _copyright;
        }

        private void bClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void url_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(urlPath);
        }
    }
}