using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Contacts;

namespace GoContactSyncMod
{
    internal interface IConflictResolver
    {
        /// <summary>
        /// Resolves contact sync conflics.
        /// </summary>
        /// <param name="outlookContact"></param>
        /// <param name="googleContact"></param>
        /// <returns>Returns ConflictResolution (enum)</returns>
        ConflictResolution Resolve(Outlook.ContactItem outlookContact, Contact googleContact);
    }

    internal enum ConflictResolution
    {
        Skip,
        Cancel,
        OutlookWins,
        GoogleWins,
        OutlookWinsAlways,
        GoogleWinsAlways,
        SkipAlways
    }

    internal enum DeleteResolution
    {
        Skip,
        Cancel,
        DeleteOutlook,
        DeleteGoogle,
        KeepOutlook,
        KeepGoogle,
        DeleteOutlookAlways,
        DeleteGoogleAlways,
        KeepOutlookAlways,
        KeepGoogleAlways,
        SkipAlways
    }
}
