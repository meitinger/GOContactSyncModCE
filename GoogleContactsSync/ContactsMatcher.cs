using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Client;
using Google.GData.Contacts;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;

namespace WebGear.GoogleContactsSync
{
    internal static class ContactsMatcher
    {
        /// <summary>
        /// Time tolerance in seconds - used when comparing date modified.
        /// </summary>
        public static int TimeTolerance = 20;

        /// <summary>
        /// Matches outlook and google contact by a) google id b) properties.
        /// </summary>
        /// <param name="outlookContacts"></param>
        /// <param name="googleContacts"></param>
        /// <returns>Returns a list of match pairs (outlook contact + google contact) for all contact. Those that weren't matche will have it's peer set to null</returns>
        public static ContactMatchList MatchContacts(Syncronizer sync)
        {
            ContactMatchList result = new ContactMatchList(Math.Max(sync.OutlookContacts.Count, sync.GoogleContacts.Capacity));
            int googleContactsMatched = 0;
            bool listingDuplicates = false;
            string duplicatesList = "";

            //for each outlook contact try to get google contact id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without google contact. 
            //foreach (Outlook._ContactItem olc in outlookContacts)
            for (int i = 1; i <= sync.OutlookContacts.Count; i++)
            {
                try
                {
                    if (!(sync.OutlookContacts[i] is Outlook.ContactItem))
                        continue;
                }
                catch (Exception ex)
                {
                    //this is needed because some contacts throw exceptions
                    continue;
                }

                Outlook.ContactItem olc = sync.OutlookContacts[i] as Outlook.ContactItem;

                // check if a duplicate
                Collection<Outlook.ContactItem> duplicates;
                if (!string.IsNullOrEmpty(olc.Email1Address))
                {
                    duplicates = sync.OutlookContactByEmail(olc.Email1Address);
                    if (duplicates.Count > 1)
                    {
                        if (!listingDuplicates)
                        {
                            duplicatesList = "Outlook contacts with the same email have been found. Please delete duplicates of:";
                            listingDuplicates = true;
                        }
                        string str = olc.FileAs + " (" + olc.Email1Address + ")";
                        if (!duplicatesList.Contains(str))
                            duplicatesList += Environment.NewLine + str;
                        continue;
                    }
                    else
                    {
                        ContactMatch dup = result.Find(delegate(ContactMatch match)
                        {
                            return match.OutlookContact != null && match.OutlookContact.Email1Address == olc.Email1Address;
                        });
                        if (dup != null)
                        {
                            if (sync.Logger != null)
                                sync.Logger.Log(string.Format("Duplicate contact found ({0}). Skipping", olc.FileAs), EventType.Information);
                            continue;
                        }
                    }
                }
                //if (duplicates.Count != 1)
                //{
                //    // if this is the first item in the duplicates
                //    // collection, then proceed, otherwise skip.
                //    int index = sync.IndexOf(duplicates, olc); //duplicates.IndexOf(olc);
                //    if (index == -1)
                //        throw new Exception("Did not find self in duplicates");
                //    if (index != 0)
                //        continue;
                //}

                if (!IsContactValid(olc))
                {
                    if (sync.Logger != null)
                        sync.Logger.Log(string.Format("Invalid outlook contact ({0}). Skipping", olc.FileAs), EventType.Warning);
                    continue;
                }

                // only match if there is either an email or telephone or else
                // a matching google contact will be created at each sync
                if (olc.Email1Address != null ||
                    olc.Email2Address != null ||
                    olc.Email3Address != null ||
                    olc.PrimaryTelephoneNumber != null || 
                    olc.HomeTelephoneNumber != null || 
                    olc.MobileTelephoneNumber != null || 
                    olc.BusinessTelephoneNumber != null
                    )
                {
                    //create a default match pair with just outlook contact.
                    ContactMatch match = new ContactMatch(olc, null);
                    #region Match by google id
                    //try to match this contact to one of google contacts.
                    Outlook.UserProperty idProp = olc.UserProperties[sync.OutlookPropertyNameId];
                    if (idProp != null)
                    {
                        AtomId id = new AtomId((string)idProp.Value);
                        ContactEntry foundContact = sync.GoogleContacts.FindById(id) as ContactEntry;

                        if (foundContact != null && foundContact.Deleted)
                        {
                            //google contact was deleted, but outlook contact is still referencing it.
                            idProp.Value = "";

                            //TODO: delete outlook contact too?
                        }
                        if (foundContact == null)
                        {
                            //google contact not found. delete property.
                            idProp.Value = "";
                        }
                        else if (foundContact != null)
                            //we found a match by google id
                            match.AddGoogleContact(foundContact);
                    }
                    #endregion

                    if (match.GoogleContact == null)
                    {
                        //no match found. match by common properties

                        #region Match by properties

                        //foreach google contac try to match and create a match pair if found some match(es)
                        foreach (ContactEntry entry in sync.GoogleContacts)
                        {
                            //Console.WriteLine(" - "+entry.Title.Text);
                            if (entry.Deleted)
                                continue;

                            //1. try to match by name
                            if (!string.IsNullOrEmpty(olc.FullName) && olc.FullName.Equals(entry.Title.Text, StringComparison.InvariantCultureIgnoreCase))
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            //1.1 try to match by file as
                            if (!string.IsNullOrEmpty(olc.FileAs) && olc.FileAs.Equals(entry.Title.Text, StringComparison.InvariantCultureIgnoreCase))
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            //2. try to match by emails
                            if (FindEmail(olc.Email1Address, entry.Emails) != null)
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            if (FindEmail(olc.Email2Address, entry.Emails) != null)
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            if (FindEmail(olc.Email3Address, entry.Emails) != null)
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            #region Phone numbers
                            //3. try to match by phone numbers
                            if (FindPhone(olc.MobileTelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            //don't match by home or business bumbers, because several people may share the saem home or business number
                            continue;

                            //if (FindPhone(olc.PrimaryTelephoneNumber, entry.Phonenumbers) != null)
                            //{
                            //    match.AddGoogleContact(entry);
                            //    continue;
                            //}                            
                            

                            if (FindPhone(olc.HomeTelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            if (FindPhone(olc.BusinessTelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                continue;
                            }

                            if (FindPhone(olc.BusinessFaxNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            }

                            if (FindPhone(olc.HomeFaxNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            }

                            if (FindPhone(olc.PagerNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            }

                            if (FindPhone(olc.RadioTelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            }

                            if (FindPhone(olc.OtherTelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            }

                            if (FindPhone(olc.CarTelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            }

                            if (FindPhone(olc.Business2TelephoneNumber, entry.Phonenumbers) != null)
                            {
                                match.AddGoogleContact(entry);
                                //continue;
                            } 
                            #endregion
                        }
                        #endregion
                    }

                    //check if a match was found.
                    if (match.AllGoogleContactMatches.Count > 0)
                    {
                        googleContactsMatched += match.AllGoogleContactMatches.Count;

                        //remove google contact from the list so it's not matched twice.
                        foreach (ContactEntry contact in match.AllGoogleContactMatches)
                            sync.GoogleContacts.Remove(contact);
                    }
                    else
                    {
                        if (sync.Logger != null)
                            sync.Logger.Log(string.Format("No match found for outlook contact ({0})", olc.FileAs), EventType.Information);
                    }

                    result.Add(match);
                }
                else
                {
                    // no telephone and email
                    if (sync.Logger != null)
                        sync.Logger.Log(string.Format("Skipping outlook contact ({0})", olc.FileAs), EventType.Warning);
                }
            }

            if (listingDuplicates)
            {
                throw new DuplicateDataException(duplicatesList);
            }

            //return result;

            if (sync.SyncOption != SyncOption.OutlookToGoogleOnly)
            {
                //for each google contact that's left (they will be nonmatched) create a new match pair without outlook contact. 
                foreach (ContactEntry entry in sync.GoogleContacts)
                {
                    // only match if there is either an email or telephone or else
                    // a matching google contact will be created at each sync
                    if (entry.Emails.Count != 0 || entry.Phonenumbers.Count != 0)
                    {
                        ContactMatch match = new ContactMatch(null, entry); ;
                        result.Add(match);
                    }
                    else
                    {
                        // no telephone and email
                    }
                }
            }
            return result;
        }

        private static bool IsContactValid(Outlook.ContactItem contact)
        {
            /*if (!string.IsNullOrEmpty(contact.FileAs))
                return true;*/

            if (!string.IsNullOrEmpty(contact.Email1Address))
                return true;

            if (!string.IsNullOrEmpty(contact.Email2Address))
                return true;

            if (!string.IsNullOrEmpty(contact.Email3Address))
                return true;

            if (!string.IsNullOrEmpty(contact.HomeTelephoneNumber))
                return true;

            if (!string.IsNullOrEmpty(contact.BusinessTelephoneNumber))
                return true;

            if (!string.IsNullOrEmpty(contact.MobileTelephoneNumber))
                return true;

            if (!string.IsNullOrEmpty(contact.HomeAddress))
                return true;

            if (!string.IsNullOrEmpty(contact.BusinessAddress))
                return true;

            if (!string.IsNullOrEmpty(contact.OtherAddress))
                return true;

            if (!string.IsNullOrEmpty(contact.Body))
                return true;

            return false;
        }

        public static void SyncContacts(Syncronizer sync)
        {
            foreach (ContactMatch match in sync.Contacts)
            {
                SyncContact(match, sync);
            }
        }
        public static void SyncContact(ContactMatch match, Syncronizer sync)
        {
            if (match.GoogleContact == null && match.OutlookContact != null)
            {
                //no google contact

                //TODO: check SyncOption
                //TODO: found that when a contacts doesn't have anything other that the name - it's not returned in the google contacts list.
                Outlook.UserProperty idProp = match.OutlookContact.UserProperties[sync.OutlookPropertyNameId];
                if (idProp != null && (string)idProp.Value!="")
                {
                    AtomId id = new AtomId((string)idProp.Value);
                    ContactEntry matchingGoogleContact = sync.GoogleContacts.FindById(id) as ContactEntry;
                    if (matchingGoogleContact == null)
                    {
                        //TODO: make sure that outlook contacts don't get deleted when deleting corresponding google contact when testing. 
                        //solution: use ResetMatching() method to unlink this relation
                        //sync.ResetMatches();
                        return;
                    }
                }

                //create a Google contact from Outlook contact
                match.GoogleContact = new ContactEntry();

                ContactSync.UpdateContact(match.OutlookContact, match.GoogleContact);
                sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);

            }
            else if (match.OutlookContact == null && match.GoogleContact != null)
            {
                // no outlook contact
                string outlookId = ContactPropertiesUtils.GetGoogleOutlookContactId(sync.SyncProfile, match.GoogleContact);
                if (outlookId != null)
                {
                    //TODO: make sure that google contacts don't get deleted when deleting corresponding outlook contact when testing. 
                    //solution: use ResetMatching() method to unlink this relation
                    //sync.ResetMatches();
                    return;
                }

                //TODO: check SyncOption
                //create a Outlook contact from Google contact
                match.OutlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;

                ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
            }
            else if (match.OutlookContact != null && match.GoogleContact != null)
            {
                //merge contact details


                //TODO: check if there are multiple matches
                if (match.AllGoogleContactMatches.Count > 1)
                {
                    //loop from 2-nd item
                    for (int m = 1; m < match.AllGoogleContactMatches.Count; m++)
                    {
                        ContactEntry entry = match.AllGoogleContactMatches[m];
                        try
                        {
                            Outlook.ContactItem item = sync.OutlookContacts.Find("[" + sync.OutlookPropertyNameId + "] = \"" + entry.Id.Uri.Content + "\"") as Outlook.ContactItem;
                            //Outlook.ContactItem item = sync.OutlookContacts.Find("[myTest] = \"value\"") as Outlook.ContactItem;
                            if (item != null)
                            {
                                //do something
                            }
                        }
                        catch (Exception)
                        {
                            //TODO: should not get here.
                        }
                    }

                    //TODO: add info to Outlook contact from extra Google contacts before deleting extra Google contacts.

                    for (int m = 1; m < match.AllGoogleContactMatches.Count; m++)
                    {
                        match.AllGoogleContactMatches[m].Delete();
                    }
                }

                //determine if this contact pair were syncronized
                //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(match.OutlookContact, sync.OutlookPropertyNameUpdated);
                DateTime? lastSynced = GetOutlookPropertyValueDateTime(match.OutlookContact, sync.OutlookPropertyNameSynced);
                if (lastSynced.HasValue)
                {
                    //contact pair was syncronysed before.

                    //determine if google contact was updated since last sync

                    //lastSynced is stored without seconds. take that into account.
                    DateTime lastUpdatedOutlook = match.OutlookContact.LastModificationTime.AddSeconds(-match.OutlookContact.LastModificationTime.Second);
                    DateTime lastUpdatedGoogle = match.GoogleContact.Updated.AddSeconds(-match.GoogleContact.Updated.Second);

                    //check if both outlok and google contacts where updated sync last sync
                    if (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds >= TimeTolerance
                        && lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds >= TimeTolerance)
                    {
                        //both contacts were updated.
                        //options: 1) ignore 2) loose one based on SyncOption
                        //throw new Exception("Both contacts were updated!");

                        switch (sync.SyncOption)
                        {
                            case SyncOption.MergeOutlookWins:
                                //overwrite google contact
                                ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                                sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);
                                break;
                            case SyncOption.MergeGoogleWins:
                                //overwrite outlook contact
                                ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                                sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                                break;
                            case SyncOption.MergePrompt:
                                //promp for sync option
                                ConflictResolver r = new ConflictResolver();
                                ConflictResolution res = r.Resolve(match.OutlookContact, match.GoogleContact);
                                switch (res)
                                {
                                    case ConflictResolution.Cancel:
                                        break;
                                    case ConflictResolution.OutlookWins:
                                        //TODO: what about categories/groups?
                                        ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                                        sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);
                                        break;
                                    case ConflictResolution.GoogleWins:
                                        //TODO: what about categories/groups?
                                        ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                                        sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            case SyncOption.GoogleToOutlookOnly:
                                ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                                sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                                break;
                            case SyncOption.OutlookToGoogleOnly:
                                ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                                sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);
                                break;
                        }
                        return;
                    }

                    //check if outlook contact was updated (with X second tolerance)
                    if (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds >= TimeTolerance)
                    {
                        //outlook contact was changed

                        //merge contacts.
                        if (sync.SyncOption != SyncOption.GoogleToOutlookOnly)
                        {
                            //TODO: use UpdateContact instead?
                            ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                            sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);

                            //at the moment use outlook as "master" source of contacts - in the event of a conflict google contact will be overwritten.
                            //TODO: control conflict resolution by SyncOption
                            return;
                        }
                    }

                    //check if google contact was updated (with X second tolerance)
                    if (lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds >= TimeTolerance)
                    {
                        //google contact was updated
                        //update outlook contact
                        if (sync.SyncOption != SyncOption.OutlookToGoogleOnly)
                        {
                            ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                            sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                        }
                    }
                }
                else
                {
                    //contacts were never synced.
                    //merge contacts.
                    switch (sync.SyncOption)
                    {
                        case SyncOption.MergeOutlookWins:
                            //overwrite google contact
                            ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                            sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);
                            break;
                        case SyncOption.MergeGoogleWins:
                            //overwrite outlook contact
                            ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                            sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                            break;
                        case SyncOption.MergePrompt:
                            //promp for sync option
                            ConflictResolver r = new ConflictResolver();
                            ConflictResolution res = r.Resolve(match.OutlookContact, match.GoogleContact);
                            switch (res)
                            {
                                case ConflictResolution.Cancel:
                                    break;
                                case ConflictResolution.OutlookWins:
                                    ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                                    sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);
                                    break;
                                case ConflictResolution.GoogleWins:
                                    ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                                    sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                                    break;
                                default:
                                    break;
                            }
                            break;
                        case SyncOption.GoogleToOutlookOnly:
                            ContactSync.MergeContacts(match.GoogleContact, match.OutlookContact);
                            sync.OverwriteContactGroups(match.GoogleContact, match.OutlookContact);
                            break;
                        case SyncOption.OutlookToGoogleOnly:
                            ContactSync.MergeContacts(match.OutlookContact, match.GoogleContact);
                            sync.OverwriteContactGroups(match.OutlookContact, match.GoogleContact);
                            break;
                    }
                }

            }
            else
                throw new ArgumentNullException("ContactMatch has all peers null.");
        }

        private static PhoneNumber FindPhone(string number, PhonenumberCollection phones)
        {
            if (string.IsNullOrEmpty(number))
                return null;

            foreach (PhoneNumber phone in phones)
            {
                if (phone.Value.Equals(number, StringComparison.InvariantCultureIgnoreCase))
                {
                    return phone;
                }
            }

            return null;
        }

        private static EMail FindEmail(string address, EMailCollection emails)
        {
            if (string.IsNullOrEmpty(address))
                return null;

            foreach (EMail email in emails)
            {
                if (address.Equals(email.Address, StringComparison.InvariantCultureIgnoreCase))
                {
                    return email;
                }
            }

            return null;
        }

        public static DateTime? GetOutlookPropertyValueDateTime(Outlook.ContactItem outlookContact, string propertyName)
        {
            Outlook.UserProperty prop = outlookContact.UserProperties[propertyName];
            if (prop != null)
                return (DateTime)prop.Value;
            return null;
        }

        /// <summary>
        /// Adds new Google Groups to the Google account.
        /// </summary>
        /// <param name="sync"></param>
        public static void SyncGroups(Syncronizer sync)
        {
            foreach (ContactMatch match in sync.Contacts)
            {
                if (match.OutlookContact != null && !string.IsNullOrEmpty(match.OutlookContact.Categories))
                {
                    string[] cats = Utilities.GetOutlookGroups(match.OutlookContact);
                    GroupEntry g;
                    foreach (string cat in cats)
                    {
                        g = sync.GetGoogleGroupByName(cat);
                        if (g == null)
                        {
                            // create group
                            if (g == null)
                            {
                                g = sync.CreateGroup(cat);
                                g = sync.SaveGoogleGroup(g);
                                sync.GoogleGroups.Add(g);
                            }
                        }
                    }
                }
            }
        }
    }

    internal class ContactMatchList : List<ContactMatch>
    {
        public ContactMatchList(int capacity) : base(capacity) { }
    }

    internal class ContactMatch
    {
        public Outlook.ContactItem OutlookContact;
        public ContactEntry GoogleContact;
        public readonly List<ContactEntry> AllGoogleContactMatches = new List<ContactEntry>(1);
        public ContactEntry LastGoogleContact;

        public ContactMatch(Outlook.ContactItem outlookContact, ContactEntry googleContact)
        {
            OutlookContact = outlookContact;
            GoogleContact = googleContact;
        }

        public void AddGoogleContact(ContactEntry googleContact)
        {
            if (googleContact == null)
                return;
            //throw new ArgumentNullException("googleContact must not be null.");

            if (GoogleContact == null)
                GoogleContact = googleContact;

            //this to avoid searching the entire collection. 
            //if last contact it what we are trying to add the we have already added it earlier
            if (LastGoogleContact == googleContact)
                return;

            if (!AllGoogleContactMatches.Contains(googleContact))
                AllGoogleContactMatches.Add(googleContact);

            LastGoogleContact = googleContact;
        }

        public void Delete()
        {
            if (GoogleContact != null)
                GoogleContact.Delete();
            if (OutlookContact != null)
                OutlookContact.Delete();
        }
    }

    //public class GroupMatchList : List<GroupMatch>
    //{
    //    public GroupMatchList(int capacity) : base(capacity) { }
    //}

    //public class GroupMatch
    //{
    //    public string OutlookGroup;
    //    public GroupEntry GoogleGroup;
    //    public readonly List<GroupEntry> AllGoogleGroupMatches = new List<GroupEntry>(1);
    //    public GroupEntry LastGoogleGroup;

    //    public GroupMatch(string outlookGroup, GroupEntry googleGroup)
    //    {
    //        OutlookGroup = outlookGroup;
    //        GoogleGroup = googleGroup;
    //    }

    //    public void AddGoogleGroup(GroupEntry googleGroup)
    //    {
    //        if (googleGroup == null)
    //            return;
    //        //throw new ArgumentNullException("googleContact must not be null.");

    //        if (GoogleGroup == null)
    //            GoogleGroup = googleGroup;

    //        //this to avoid searching the entire collection. 
    //        //if last contact it what we are trying to add the we have already added it earlier
    //        if (LastGoogleGroup == googleGroup)
    //            return;

    //        if (!AllGoogleGroupMatches.Contains(googleGroup))
    //            AllGoogleGroupMatches.Add(googleGroup);

    //        LastGoogleGroup = googleGroup;
    //    }
    //}
}
