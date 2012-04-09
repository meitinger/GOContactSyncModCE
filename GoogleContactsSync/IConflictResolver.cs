using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Contacts;
using Google.Documents;

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
        ConflictResolution Resolve(ContactMatch match, bool isNewMatch);

        ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.NoteItem outlookNote, Document googleNote, Syncronizer sync, bool isNewMatch);

        ConflictResolution ResolveDuplicate(OutlookContactInfo outlookContact, List<Contact> googleContacts, out Contact googleContact);

        DeleteResolution ResolveDelete(OutlookContactInfo outlookContact);

        DeleteResolution ResolveDelete(Contact googleContact);

        DeleteResolution ResolveDelete(Document googleNote, Syncronizer sync);

        DeleteResolution ResolveDelete(Microsoft.Office.Interop.Outlook.NoteItem outlookNote);

        
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
        Cancel,
        DeleteOutlook,
        DeleteGoogle,
        KeepOutlook,
        KeepGoogle,
        DeleteOutlookAlways,
        DeleteGoogleAlways,
        KeepOutlookAlways,
        KeepGoogleAlways
    }    
}
