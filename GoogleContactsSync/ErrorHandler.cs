using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace WebGear.GoogleContactsSync
{
    static class ErrorHandler
    {
        public static void Handle(Exception ex)
        {
            // TODO: Write a nice error dialog, that maybe supports directly E-Mail sending as bugreport

            string message = "Sorry, an unexpected error occured.\nPlease support us fixing this problem. Go to\nhttps://sourceforge.net/projects/googlesyncmod/ and use the Tracker.\nHint: You can copy this message by pressing CTRL-C in the dialog box.\n\nError Details:\n{0}";
            message = string.Format(message, ex.ToString());
            MessageBox.Show(message, "Google Contact Sync");
        }
    }
}
