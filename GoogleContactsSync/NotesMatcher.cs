using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Documents;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    internal static class NotesMatcher
    {
        /// <summary>
        /// Time tolerance in seconds - used when comparing date modified.
        /// Less than 60 seconds doesn't make sense, as the lastSync is saved without seconds and if it is compared
        /// with the LastUpdate dates of Google and Outlook, in the worst case you compare e.g. 15:59 with 16:00 and 
        /// after truncating to minutes you compare 15:00 wiht 16:00
        /// </summary>
        public static int TimeTolerance = 60;

        public delegate void NotificationHandler(string message);
        public static event NotificationHandler NotificationReceived;

        /// <summary>
        /// Matches outlook and google note by a) google id b) properties.
        /// </summary>
        /// <param name="outlookNotes"></param>
        /// <param name="googleNotes"></param>
        /// <returns>Returns a list of match pairs (outlook note + google note) for all note. Those that weren't matche will have it's peer set to null</returns>
        public static List<NoteMatch> MatchNotes(Syncronizer sync)
        {
            Logger.Log("Matching Outlook and Google notes...", EventType.Information);
            List<NoteMatch> result = new List<NoteMatch>();

            //string duplicateGoogleMatches = "";
            //string duplicateOutlookNotes = "";
            //sync.GoogleNoteDuplicates = new Collection<NoteMatch>();
            //sync.OutlookNoteDuplicates = new Collection<NoteMatch>();

            List<string> skippedOutlookIds = new List<string>();

            //for each outlook note try to get google note id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without google note. 
            //foreach (Outlook._NoteItem olc in outlookNotes)
            Collection<Outlook.NoteItem> outlookNotesWithoutOutlookGoogleId = new Collection<Outlook.NoteItem>();
            #region Match first all outlookNotes by sync id
            for (int i = 1; i <= sync.OutlookNotes.Count; i++)
            {
                Outlook.NoteItem oln = null;
                try
                {
                    oln = sync.OutlookNotes[i] as Outlook.NoteItem;
                    if (oln == null)
                    {
                        Logger.Log("Empty Outlook note found. Skipping", EventType.Warning);
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    //this is needed because some notes throw exceptions
                    Logger.Log("Accessing Outlook note threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                //try
                //{
                    
                    if (NotificationReceived != null)
                        NotificationReceived(String.Format("Matching note {0} of {1} by id: {2} ...", i, sync.OutlookNotes.Count, oln.Subject));

                    // Create our own info object to go into collections/lists, so we can free the Outlook objects and not run out of resources / exceed policy limits.
                    //OutlookNoteInfo olci = new OutlookNoteInfo(oln, sync);

                    //try to match this note to one of google notes
                    Outlook.ItemProperties userProperties = oln.ItemProperties;
                    Outlook.ItemProperty idProp = userProperties[sync.OutlookPropertyNameId];
                    try
                    {
                        if (idProp != null)
                        {
                            string googleNoteId = string.Copy((string)idProp.Value);
                            Document foundNote = sync.GetGoogleNoteById(googleNoteId);
                            NoteMatch match = new NoteMatch(oln, null);

                            //Check first, that this is not a duplicate 
                            //e.g. by copying an existing Outlook note
                            //or by Outlook checked this as duplicate, but the user selected "Add new"
                        //    Collection<OutlookNoteInfo> duplicates = sync.OutlookNoteByProperty(sync.OutlookPropertyNameId, googleNoteId);
                        //    if (duplicates.Count > 1)
                        //    {
                        //        foreach (OutlookNoteInfo duplicate in duplicates)
                        //        {
                        //            if (!string.IsNullOrEmpty(googleNoteId))
                        //            {
                        //                Logger.Log("Duplicate Outlook note found, resetting match and try to match again: " + duplicate.FileAs, EventType.Warning);
                        //                idProp.Value = "";
                        //            }
                        //        }

                        //        if (foundNote != null && !foundNote.Deleted)
                        //        {
                        //            NotePropertiesUtils.ResetGoogleOutlookNoteId(sync.SyncProfile, foundNote);
                        //        }

                        //        outlookNotesWithoutOutlookGoogleId.Add(olci);
                        //    }
                        //    else
                        //    {

                            if (foundNote != null)
                            {
                                //we found a match by google id, that is not deleted yet
                                match.AddGoogleNote(foundNote);
                                result.Add(match);
                                //Remove the note from the list to not sync it twice
                                sync.GoogleNotes.Remove(foundNote);
                            }
                            else
                            {
                                ////If no match found, is the note either deleted on Google side or was a copy on Outlook side 
                                ////If it is a copy on Outlook side, the idProp.Value must be emptied to assure, the note is created on Google side and not deleted on Outlook side
                                ////bool matchIsDuplicate = false;
                                //foreach (NoteMatch existingMatch in result)
                                //{
                                //    if (existingMatch.OutlookNote.UserProperties[sync.OutlookPropertyNameId].Value.Equals(idProp.Value))
                                //    {
                                //        //matchIsDuplicate = true;
                                //        idProp.Value = "";
                                //        break;
                                //    }

                                //}
                                outlookNotesWithoutOutlookGoogleId.Add(oln);

                                //if (!matchIsDuplicate)
                                //    result.Add(match);
                            }
                        //    }
                        }
                        else
                            outlookNotesWithoutOutlookGoogleId.Add(oln);
                    }
                    finally
                    {
                        if (idProp != null)
                            Marshal.ReleaseComObject(idProp);
                        Marshal.ReleaseComObject(userProperties);
                    }
                //}

                //finally
                //{
                //    Marshal.ReleaseComObject(oln);
                //    oln = null;
                //}

            }
            #endregion
            #region Match the remaining notes by properties

            for (int i = 0; i < outlookNotesWithoutOutlookGoogleId.Count; i++)
            {
                Outlook.NoteItem oln = outlookNotesWithoutOutlookGoogleId[i];

                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Matching note {0} of {1} by unique properties: {2} ...", i + 1, outlookNotesWithoutOutlookGoogleId.Count, oln.Subject));

                //no match found by id => match by subject/title
                //create a default match pair with just outlook note.
                NoteMatch match = new NoteMatch(oln, null);

                //foreach google contact try to match and create a match pair if found some match(es)
                for (int j = sync.GoogleNotes.Count - 1; j >= 0; j--)
                {
                    Document entry = sync.GoogleNotes[j];

                    string body = NotePropertiesUtils.GetBody(sync, entry);
                    if (!String.IsNullOrEmpty(body))
                        body = body.Replace("\r\n", string.Empty).Replace(" ", string.Empty).Replace(" ", string.Empty);

                    string outlookBody = null;
                    if (!string.IsNullOrEmpty(oln.Body))
                        outlookBody = oln.Body.Replace("\t", "        ").Replace("\r\n", string.Empty).Replace(" ", string.Empty).Replace(" ", string.Empty);

                    // only match if there is a note body, else
                    // a matching google note will be created at each sync                
                    if (!string.IsNullOrEmpty(outlookBody) && !string.IsNullOrEmpty(body) && outlookBody.Equals(body, StringComparison.InvariantCultureIgnoreCase)
                        )
                    {
                        match.AddGoogleNote(entry);
                        sync.GoogleNotes.Remove(entry);
                    }

                }

                #region find duplicates not needed now
                //if (match.GoogleNote == null && match.OutlookNote != null)
                //{//If GoogleNote, we have to expect a conflict because of Google insert of duplicates
                //    foreach (Note entry in sync.GoogleNotes)
                //    {                        
                //        if (!string.IsNullOrEmpty(olc.FullName) && olc.FullName.Equals(entry.Title, StringComparison.InvariantCultureIgnoreCase) ||
                //         !string.IsNullOrEmpty(olc.FileAs) && olc.FileAs.Equals(entry.Title, StringComparison.InvariantCultureIgnoreCase) ||
                //         !string.IsNullOrEmpty(olc.Email1Address) && FindEmail(olc.Email1Address, entry.Emails) != null ||
                //         !string.IsNullOrEmpty(olc.Email2Address) && FindEmail(olc.Email1Address, entry.Emails) != null ||
                //         !string.IsNullOrEmpty(olc.Email3Address) && FindEmail(olc.Email1Address, entry.Emails) != null ||
                //         olc.MobileTelephoneNumber != null && FindPhone(olc.MobileTelephoneNumber, entry.Phonenumbers) != null
                //         )
                //    }
                //// check for each email 1,2 and 3 if a duplicate exists with same email, because Google doesn't like inserting new notes with same email
                //Collection<Outlook.NoteItem> duplicates1 = new Collection<Outlook.NoteItem>();
                //Collection<Outlook.NoteItem> duplicates2 = new Collection<Outlook.NoteItem>();
                //Collection<Outlook.NoteItem> duplicates3 = new Collection<Outlook.NoteItem>();
                //if (!string.IsNullOrEmpty(olc.Email1Address))
                //    duplicates1 = sync.OutlookNoteByEmail(olc.Email1Address);

                //if (!string.IsNullOrEmpty(olc.Email2Address))
                //    duplicates2 = sync.OutlookNoteByEmail(olc.Email2Address);

                //if (!string.IsNullOrEmpty(olc.Email3Address))
                //    duplicates3 = sync.OutlookNoteByEmail(olc.Email3Address);


                //if (duplicates1.Count > 1 || duplicates2.Count > 1 || duplicates3.Count > 1)
                //{
                //    if (string.IsNullOrEmpty(duplicatesEmailList))
                //        duplicatesEmailList = "Outlook notes with the same email have been found and cannot be synchronized. Please delete duplicates of:";

                //    if (duplicates1.Count > 1)
                //        foreach (Outlook.NoteItem duplicate in duplicates1)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email1Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    if (duplicates2.Count > 1)
                //        foreach (Outlook.NoteItem duplicate in duplicates2)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email2Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    if (duplicates3.Count > 1)
                //        foreach (Outlook.NoteItem duplicate in duplicates3)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email3Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    continue;
                //}
                //else if (!string.IsNullOrEmpty(olc.Email1Address))
                //{
                //    NoteMatch dup = result.Find(delegate(NoteMatch match)
                //    {
                //        return match.OutlookNote != null && match.OutlookNote.Email1Address == olc.Email1Address;
                //    });
                //    if (dup != null)
                //    {
                //        Logger.Log(string.Format("Duplicate note found by Email1Address ({0}). Skipping", olc.FileAs), EventType.Information);
                //        continue;
                //    }
                //}

                //// check for unique mobile phone, because this sync tool uses the also the mobile phone to identify matches between Google and Outlook
                //Collection<Outlook.NoteItem> duplicatesMobile = new Collection<Outlook.NoteItem>();
                //if (!string.IsNullOrEmpty(olc.MobileTelephoneNumber))
                //    duplicatesMobile = sync.OutlookNoteByProperty("MobileTelephoneNumber", olc.MobileTelephoneNumber);

                //if (duplicatesMobile.Count > 1)
                //{
                //    if (string.IsNullOrEmpty(duplicatesMobileList))
                //        duplicatesMobileList = "Outlook notes with the same mobile phone have been found and cannot be synchronized. Please delete duplicates of:";

                //    foreach (Outlook.NoteItem duplicate in duplicatesMobile)
                //    {
                //        sync.OutlookNoteDuplicates.Add(olc);
                //        string str = olc.FileAs + " (" + olc.MobileTelephoneNumber + ")";
                //        if (!duplicatesMobileList.Contains(str))
                //            duplicatesMobileList += Environment.NewLine + str;
                //    }
                //    continue;
                //}
                //else if (!string.IsNullOrEmpty(olc.MobileTelephoneNumber))
                //{
                //    NoteMatch dup = result.Find(delegate(NoteMatch match)
                //    {
                //        return match.OutlookNote != null && match.OutlookNote.MobileTelephoneNumber == olc.MobileTelephoneNumber;
                //    });
                //    if (dup != null)
                //    {
                //        Logger.Log(string.Format("Duplicate note found by MobileTelephoneNumber ({0}). Skipping", olc.FileAs), EventType.Information);
                //        continue;
                //    }
                //}

                #endregion

            //    if (match.AllGoogleNoteMatches == null || match.AllGoogleNoteMatches.Count == 0)
            //    {
            //        //Check, if this Outlook note has a match in the google duplicates
            //        bool duplicateFound = false;
            //        foreach (NoteMatch duplicate in sync.GoogleNoteDuplicates)
            //        {
            //            if (duplicate.AllGoogleNoteMatches.Count > 0 &&
            //                (!string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(duplicate.AllGoogleNoteMatches[0].Title) && olci.FileAs.Equals(duplicate.AllGoogleNoteMatches[0].Title.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||  //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google
            //                 !string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(duplicate.AllGoogleNoteMatches[0].Name.FullName) && olci.FileAs.Equals(duplicate.AllGoogleNoteMatches[0].Name.FullName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
            //                 !string.IsNullOrEmpty(olci.FullName) && !string.IsNullOrEmpty(duplicate.AllGoogleNoteMatches[0].Name.FullName) && olci.FullName.Equals(duplicate.AllGoogleNoteMatches[0].Name.FullName.Replace("\r\n", "\n").Replace("\n", "\r\n"), StringComparison.InvariantCultureIgnoreCase) ||
            //                 !string.IsNullOrEmpty(olci.Email1Address) && duplicate.AllGoogleNoteMatches[0].Emails.Count > 0 && olci.Email1Address.Equals(duplicate.AllGoogleNoteMatches[0].Emails[0].Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //!string.IsNullOrEmpty(olci.Email2Address) && FindEmail(olci.Email2Address, duplicate.AllGoogleNoteMatches[0].Emails) != null ||
            //                //!string.IsNullOrEmpty(olci.Email3Address) && FindEmail(olci.Email3Address, duplicate.AllGoogleNoteMatches[0].Emails) != null ||
            //                 olci.MobileTelephoneNumber != null && FindPhone(olci.MobileTelephoneNumber, duplicate.AllGoogleNoteMatches[0].Phonenumbers) != null ||
            //                 !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(duplicate.AllGoogleNoteMatches[0].Title) && duplicate.AllGoogleNoteMatches[0].Organizations.Count > 0 && olci.FileAs.Equals(duplicate.AllGoogleNoteMatches[0].Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
            //                ) ||
            //                !string.IsNullOrEmpty(olci.FileAs) && olci.FileAs.Equals(duplicate.OutlookNote.Subject, StringComparison.InvariantCultureIgnoreCase) ||
            //                !string.IsNullOrEmpty(olci.FullName) && olci.FullName.Equals(duplicate.OutlookNote.FullName, StringComparison.InvariantCultureIgnoreCase) ||
            //                !string.IsNullOrEmpty(olci.Email1Address) && olci.Email1Address.Equals(duplicate.OutlookNote.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //                                              olci.Email1Address.Equals(duplicate.OutlookNote.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //                                              olci.Email1Address.Equals(duplicate.OutlookNote.Email3Address, StringComparison.InvariantCultureIgnoreCase)
            //                //                                              ) ||
            //                //!string.IsNullOrEmpty(olci.Email2Address) && (olci.Email2Address.Equals(duplicate.OutlookNote.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //                                              olci.Email2Address.Equals(duplicate.OutlookNote.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //                                              olci.Email2Address.Equals(duplicate.OutlookNote.Email3Address, StringComparison.InvariantCultureIgnoreCase)
            //                //                                              ) ||
            //                //!string.IsNullOrEmpty(olci.Email3Address) && (olci.Email3Address.Equals(duplicate.OutlookNote.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //                                              olci.Email3Address.Equals(duplicate.OutlookNote.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
            //                //                                              olci.Email3Address.Equals(duplicate.OutlookNote.Email3Address, StringComparison.InvariantCultureIgnoreCase)
            //                //                                              ) ||
            //                olci.MobileTelephoneNumber != null && olci.MobileTelephoneNumber.Equals(duplicate.OutlookNote.MobileTelephoneNumber) ||
            //                !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(duplicate.GoogleNote.Title) && duplicate.GoogleNote.Organizations.Count > 0 && olci.FileAs.Equals(duplicate.GoogleNote.Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
            //               )
            //            {
            //                duplicateFound = true;
            //                sync.OutlookNoteDuplicates.Add(match);
            //                if (string.IsNullOrEmpty(duplicateOutlookNotes))
            //                    duplicateOutlookNotes = "Outlook note found that has been already identified as duplicate Google note (either same email, Mobile or FullName) and cannot be synchronized. Please delete or resolve duplicates of:";

            //                string str = olci.FileAs + " (" + olci.Email1Address + ", " + olci.MobileTelephoneNumber + ")";
            //                if (!duplicateOutlookNotes.Contains(str))
            //                    duplicateOutlookNotes += Environment.NewLine + str;
            //            }
            //        }

            //        if (!duplicateFound)
            if (match.GoogleNote == null)
                Logger.Log(string.Format("No match found for outlook note ({0}) => {1}", match.OutlookNote.Subject, (NotePropertiesUtils.GetOutlookGoogleNoteId(sync, match.OutlookNote) != null ? "Delete from Outlook" : "Add to Google")), EventType.Information);
            
            //    }
            //    else
            //    {
            //        //Remember Google duplicates to later react to it when resetting matches or syncing
            //        //ResetMatches: Also reset the duplicates
            //        //Sync: Skip duplicates (don't sync duplicates to be fail safe)
            //        if (match.AllGoogleNoteMatches.Count > 1)
            //        {
            //            sync.GoogleNoteDuplicates.Add(match);
            //            foreach (Note entry in match.AllGoogleNoteMatches)
            //            {
            //                //Create message for duplicatesFound exception
            //                if (string.IsNullOrEmpty(duplicateGoogleMatches))
            //                    duplicateGoogleMatches = "Outlook notes matching with multiple Google notes have been found (either same email, Mobile, FullName or company) and cannot be synchronized. Please delete or resolve duplicates of:";

            //                string str = olci.FileAs + " (" + olci.Email1Address + ", " + olci.MobileTelephoneNumber + ")";
            //                if (!duplicateGoogleMatches.Contains(str))
            //                    duplicateGoogleMatches += Environment.NewLine + str;
            //            }
            //        }



            //    }                

                result.Add(match);
            }
            #endregion

            //if (!string.IsNullOrEmpty(duplicateGoogleMatches) || !string.IsNullOrEmpty(duplicateOutlookNotes))
            //    duplicatesFound = new DuplicateDataException(duplicateGoogleMatches + Environment.NewLine + Environment.NewLine + duplicateOutlookNotes);
            //else
            //    duplicatesFound = null;

            //return result;

            //for each google note that's left (they will be nonmatched) create a new match pair without outlook note. 
            for (int i = 0; i < sync.GoogleNotes.Count; i++)
            {
                Document entry = sync.GoogleNotes[i];               
                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Adding new Google note {0} of {1} by unique properties: {2} ...", i + 1, sync.GoogleNotes.Count, entry.Title));

                //string googleOutlookId = NotePropertiesUtils.GetGoogleOutlookNoteId(sync.SyncProfile, entry);
                //if (!String.IsNullOrEmpty(googleOutlookId) && skippedOutlookIds.Contains(googleOutlookId))
                //{
                //    Logger.Log("Skipped GoogleNote because Outlook note couldn't be matched beacause of previous problem (see log): " + entry.Title, EventType.Warning);
                //}
                //else 
                if (string.IsNullOrEmpty(entry.Title) && string.IsNullOrEmpty(entry.Content))
                {
                    // no title or content
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    Logger.Log("Skipped GoogleNote because no unique property found (Title or Content):" + entry.Title, EventType.Warning);
                }
                else
                {
                    Logger.Log(string.Format("No match found for google note ({0}) => {1}", entry.Title, (NotePropertiesUtils.NoteFileExists(entry.Id, sync.SyncProfile) ? "Delete from Google" : "Add to Outlook")), EventType.Information);
                    NoteMatch match = new NoteMatch(null, entry);
                    result.Add(match);
                }
            }
            return result;
        }

        



        public static void SyncNotes(Syncronizer sync)
        {
            for (int i = 0; i < sync.Notes.Count; i++)
            {
                NoteMatch match = sync.Notes[i];
                if (NotificationReceived != null)
                {
                    string name = string.Empty;
                    if (match.OutlookNote != null)
                        name = match.OutlookNote.Subject;
                    else if (match.GoogleNote != null)
                        name = match.GoogleNote.Title;
                    NotificationReceived(String.Format("Syncing note {0} of {1}: {2} ...", i + 1, sync.Notes.Count, name));
                }

                SyncNote(match, sync);
            }
        }
        public static void SyncNote(NoteMatch match, Syncronizer sync)
        {
            Outlook.NoteItem outlookNoteItem = match.OutlookNote;
           
            //try
            //{
                if (match.GoogleNote == null && match.OutlookNote != null)
                {
                    //no google note                               
                    string googleNotetId = NotePropertiesUtils.GetOutlookGoogleNoteId(sync, outlookNoteItem);
                    if (!string.IsNullOrEmpty(googleNotetId))
                    {                        
                        //Redundant check if exist, but in case an error occurred in MatchNotes
                        Document matchingGoogleNote = sync.GetGoogleNoteById(googleNotetId);
                        if (matchingGoogleNote == null)
                            if (!sync.PromptDelete)
                                sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlookAlways;
                            else if (sync.DeleteOutlookResolution != DeleteResolution.DeleteOutlookAlways &&
                                     sync.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                            {
                                ConflictResolver r = new ConflictResolver();
                                sync.DeleteOutlookResolution = r.ResolveDelete(match.OutlookNote);
                            }
                        switch (sync.DeleteOutlookResolution)
                        {
                            case DeleteResolution.KeepOutlook:
                            case DeleteResolution.KeepOutlookAlways:
                                NotePropertiesUtils.ResetOutlookGoogleNoteId(sync, match.OutlookNote);
                                break;
                            case DeleteResolution.DeleteOutlook:
                            case DeleteResolution.DeleteOutlookAlways:
                                //Avoid recreating a GoogleNote already existing
                                //==> Delete this outlookNote instead if previous match existed but no match exists anymore
                                return;
                            default:
                                throw new ApplicationException("Cancelled");
                        }
                    }

                    if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        sync.SkippedCount++;
                        Logger.Log(string.Format("Outlook Note not added to Google, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.OutlookNote.Subject), EventType.Information);
                        return;
                    }

                    //create a Google note from Outlook note
                    match.GoogleNote = new Document();
                    match.GoogleNote.Type = Document.DocumentType.Document;
                    //match.GoogleNote.Categories.Add(new AtomCategory("http://schemas.google.com/docs/2007#document"));
                    //match.GoogleNote.Categories.Add(new AtomCategory("document"));

                    sync.UpdateNote(outlookNoteItem, match.GoogleNote);

                }
                else if (match.OutlookNote == null && match.GoogleNote != null)
                {

                    // no outlook note
                    if (NotePropertiesUtils.NoteFileExists(match.GoogleNote.Id, sync.SyncProfile))
                    {                                       
                        if (!sync.PromptDelete)
                            sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                        else if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                                 sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                        {
                            ConflictResolver r = new ConflictResolver();
                            sync.DeleteGoogleResolution = r.ResolveDelete(match.GoogleNote, sync);
                        }
                        switch (sync.DeleteGoogleResolution)
                        {
                            case DeleteResolution.KeepGoogle:
                            case DeleteResolution.KeepGoogleAlways:
                                System.IO.File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id, sync.SyncProfile));                                
                                break;
                            case DeleteResolution.DeleteGoogle:
                            case DeleteResolution.DeleteGoogleAlways:
                                //Avoid recreating a OutlookNote already existing
                                //==> Delete this googleNote instead if previous match existed but no match exists anymore 
                                return;
                            default:
                                throw new ApplicationException("Cancelled");
                        }         
                    }


                    if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        sync.SkippedCount++;
                        Logger.Log(string.Format("Google Note not added to Outlook, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.GoogleNote.Title), EventType.Information);
                        return;
                    }

                    //create a Outlook note from Google note
                    outlookNoteItem = Syncronizer.CreateOutlookNoteItem(sync.SyncNotesFolder);

                    sync.UpdateNote(match.GoogleNote, outlookNoteItem);
                    match.OutlookNote = outlookNoteItem;
                }
                else if (match.OutlookNote != null && match.GoogleNote != null)
                {
                    //merge note details                

                    //determine if this note pair were syncronized
                    //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(match.OutlookNote, sync.OutlookPropertyNameUpdated);
                    DateTime? lastSynced = NotePropertiesUtils.GetOutlookLastSync(sync,outlookNoteItem);
                    if (lastSynced.HasValue)
                    {
                        //note pair was syncronysed before.

                        //determine if google note was updated since last sync

                        //lastSynced is stored without seconds. take that into account.
                        DateTime lastUpdatedOutlook = match.OutlookNote.LastModificationTime.AddSeconds(-match.OutlookNote.LastModificationTime.Second);
                        DateTime lastUpdatedGoogle = match.GoogleNote.Updated.AddSeconds(-match.GoogleNote.Updated.Second);

                        //check if both outlok and google notes where updated sync last sync
                        if (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance
                            && lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance)
                        {
                            //both notes were updated.
                            //options: 1) ignore 2) loose one based on SyncOption
                            //throw new Exception("Both notes were updated!");

                            switch (sync.SyncOption)
                            {
                                case SyncOption.MergeOutlookWins:
                                case SyncOption.OutlookToGoogleOnly:
                                    //overwrite google note
                                    Logger.Log("Outlook and Google note have been updated, Outlook note is overwriting Google because of SyncOption " + sync.SyncOption + ": " + match.OutlookNote.Subject + ".", EventType.Information);
                                    sync.UpdateNote(outlookNoteItem, match.GoogleNote);
                                    break;
                                case SyncOption.MergeGoogleWins:
                                case SyncOption.GoogleToOutlookOnly:
                                    //overwrite outlook note
                                    Logger.Log("Outlook and Google note have been updated, Google note is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + match.OutlookNote.Subject + ".", EventType.Information);
                                    sync.UpdateNote(match.GoogleNote, outlookNoteItem);
                                    break;
                                case SyncOption.MergePrompt:
                                    //promp for sync option
                                    if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                        sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                        sync.ConflictResolution != ConflictResolution.SkipAlways)
                                    {
                                        ConflictResolver r = new ConflictResolver();
                                        sync.ConflictResolution = r.Resolve(outlookNoteItem, match.GoogleNote, sync, false);
                                    }
                                    switch (sync.ConflictResolution)
                                    {
                                        case ConflictResolution.Skip:
                                        case ConflictResolution.SkipAlways:
                                            Logger.Log(string.Format("User skipped note ({0}).", match.ToString()), EventType.Information);
                                            sync.SkippedCount++;
                                            break;
                                        case ConflictResolution.OutlookWins:
                                        case ConflictResolution.OutlookWinsAlways:
                                            sync.UpdateNote(outlookNoteItem, match.GoogleNote);
                                            break;
                                        case ConflictResolution.GoogleWins:
                                        case ConflictResolution.GoogleWinsAlways:
                                            sync.UpdateNote(match.GoogleNote, outlookNoteItem);
                                            break;
                                        default:
                                            throw new ApplicationException("Canceled");
                                    }
                                    break;
                            }
                            return;
                        }


                        //check if outlook note was updated (with X second tolerance)
                        if (sync.SyncOption != SyncOption.GoogleToOutlookOnly &&
                            (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                             lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                             sync.SyncOption == SyncOption.OutlookToGoogleOnly
                            )
                           )
                        {
                            //outlook note was changed or changed Google note will be overwritten

                            if (lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                                sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                                Logger.Log("Google note has been updated since last sync, but Outlook note is overwriting Google because of SyncOption " + sync.SyncOption + ": " + match.OutlookNote.Subject + ".", EventType.Information);

                            sync.UpdateNote(outlookNoteItem, match.GoogleNote);

                            //at the moment use outlook as "master" source of notes - in the event of a conflict google note will be overwritten.
                            //TODO: control conflict resolution by SyncOption
                            return;
                        }

                        //check if google note was updated (with X second tolerance)
                        if (sync.SyncOption != SyncOption.OutlookToGoogleOnly &&
                            (lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                             lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                             sync.SyncOption == SyncOption.GoogleToOutlookOnly
                            )
                           )
                        {
                            //google note was changed or changed Outlook note will be overwritten

                            if (lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                                sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                                Logger.Log("Outlook note has been updated since last sync, but Google note is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + match.OutlookNote.Subject + ".", EventType.Information);

                            sync.UpdateNote(match.GoogleNote, outlookNoteItem);
                        }
                    }
                    else
                    {
                        //notes were never synced.
                        //merge notes.
                        switch (sync.SyncOption)
                        {
                            case SyncOption.MergeOutlookWins:
                            case SyncOption.OutlookToGoogleOnly:
                                //overwrite google note
                                sync.UpdateNote(outlookNoteItem, match.GoogleNote);
                                break;
                            case SyncOption.MergeGoogleWins:
                            case SyncOption.GoogleToOutlookOnly:
                                //overwrite outlook note
                                sync.UpdateNote(match.GoogleNote, outlookNoteItem);
                                break;
                            case SyncOption.MergePrompt:
                                //promp for sync option
                                if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                    sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                        sync.ConflictResolution != ConflictResolution.SkipAlways)
                                {
                                    ConflictResolver r = new ConflictResolver();
                                    sync.ConflictResolution = r.Resolve(outlookNoteItem, match.GoogleNote, sync, true);
                                }
                                switch (sync.ConflictResolution)
                                {
                                    case ConflictResolution.Skip:                                    
                                    case ConflictResolution.SkipAlways: //Keep both, Google AND Outlook
                                        sync.Notes.Add(new NoteMatch(match.OutlookNote, null));
                                        sync.Notes.Add(new NoteMatch(null, match.GoogleNote));
                                        break;
                                    case ConflictResolution.OutlookWins:
                                    case ConflictResolution.OutlookWinsAlways:
                                        sync.UpdateNote(outlookNoteItem, match.GoogleNote);
                                        break;
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways:
                                        sync.UpdateNote(match.GoogleNote, outlookNoteItem);
                                        break;
                                    default:
                                        throw new ApplicationException("Canceled");
                                }
                                break;
                        }
                    }

                }
                else
                    throw new ArgumentNullException("NotetMatch has all peers null.");
            //}
            //finally
            //{
                //if (outlookNoteItem != null &&
                //    match.OutlookNote != null)
                //{
                //    match.OutlookNote.Update(outlookNoteItem, sync);
                //    Marshal.ReleaseComObject(outlookNoteItem);
                //    outlookNoteItem = null;
                //}
            //}

        }
    }



    internal class NoteMatch
    {
        //ToDo: OutlookNoteInfo
        public Outlook.NoteItem OutlookNote;
        public Document GoogleNote;
        public readonly List<Document> AllGoogleNoteMatches = new List<Document>(1);
        public Document LastGoogleNote;
        public bool AsyncUpdateCompleted = false;

        public NoteMatch(Outlook.NoteItem outlookNote, Document googleNote)
        {
            OutlookNote = outlookNote;
            GoogleNote = googleNote;
        }

        public void AddGoogleNote(Document googleNote)
        {
            if (googleNote == null)
                return;
            //throw new ArgumentNullException("googleNote must not be null.");

            if (GoogleNote == null)
                GoogleNote = googleNote;

            //this to avoid searching the entire collection. 
            //if last note it what we are trying to add the we have already added it earlier
            if (LastGoogleNote == googleNote)
                return;

            if (!AllGoogleNoteMatches.Contains(googleNote))
                AllGoogleNoteMatches.Add(googleNote);

            LastGoogleNote = googleNote;
        }

        //public void Delete(NotesRequest googleService)
        //{
        //    if (GoogleNote != null)
        //         googleService.Delete(GoogleNote);
        //    if (OutlookNote != null)
        //        OutlookNote.Delete();
        //}
    }
    
    		
}
