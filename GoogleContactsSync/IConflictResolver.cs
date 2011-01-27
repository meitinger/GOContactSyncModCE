using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal interface IConflictResolver
    {
        /// <summary>
        /// Resolves contact sync conflics.
        /// </summary>
        /// <param name="outlookContact"></param>
        /// <param name="googleContact"></param>
        /// <returns>Returns true is a conflict was resolved. If cancel was pressed then returns false</returns>
        ConflictResolution Resolve(Outlook.ContactItem outlookContact, ContactEntry googleContact);
    }

    internal enum ConflictResolution
    {
        Skip,
        Cancel,
        OutlookWins,
        GoogleWins,
    }
}
