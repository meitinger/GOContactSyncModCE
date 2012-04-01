using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using System.Collections;
using Google.Documents;
using System.Runtime.InteropServices;
using System.IO;

namespace GoContactSyncMod
{
    internal static class NotePropertiesUtils
    {
        public static string GetOutlookId(Outlook.NoteItem outlookNote)
        {
            return outlookNote.EntryID;
        }
        public static string GetGoogleId(Document googleNote)
        {
            string id = googleNote.Id.ToString();
            if (id == null)
                throw new Exception();
            return id;
        }

        //public static void SetGoogleOutlookNoteId(string syncProfile, Document googleNote, Outlook.NoteItem outlookNote)
        //{
        //    if (outlookNote.EntryID == null)
        //        throw new Exception("Must save outlook note before getting id");

        //    SetGoogleOutlookNoteId(syncProfile, googleNote, GetOutlookId(outlookNote));
        //}
        //public static void SetGoogleOutlookNoteId(string syncProfile, Document googleNote, string outlookNoteId)
        //{
        //    // check if exists
        //    bool found = false;
        //    foreach (Google.GData.Extensions.ExtendedProperty p in googleNote.ExtendedProperties)
        //    {
        //        if (p.Name == "gos:oid:" + syncProfile + "")
        //        {
        //            p.Value = outlookNoteId;
        //            found = true;
        //            break;
        //        }
        //    }
        //    if (!found)
        //    {
        //        Google.GData.Extensions.ExtendedProperty prop = new ExtendedProperty(outlookNoteId, "gos:oid:" + syncProfile + "");
        //        prop.Value = outlookNoteId;
        //        googleNote.ExtendedProperties.Add(prop);
        //    }

        //}
        //public static string GetGoogleOutlookNoteId(string syncProfile, Document googleNote)
        //{
        //    // get extended prop
        //    foreach (Google.GData.Extensions.ExtendedProperty p in googleNote.DocumentEntry.ExtendedProperties)
        //    {
        //        if (p.Name == "gos:oid:" + syncProfile + "")
        //            return (string)p.Value;
        //    }
        //    return null;
        //}
        //public static void ResetGoogleOutlookNoteId(string syncProfile, Document googleNote)
        //{
        //    // get extended prop
        //    foreach (Google.GData.Extensions.ExtendedProperty p in googleNote.ExtendedProperties)
        //    {
        //        if (p.Name == "gos:oid:" + syncProfile + "")
        //        {
        //            // remove 
        //            googleNote.ExtendedProperties.Remove(p);
        //            return;
        //        }
        //    }
        //}

        /// <summary>
        /// Sets the syncId of the Outlook note and the last sync date. 
        /// Please assure to always call this function when saving OutlookItem
        /// </summary>
        /// <param name="sync"></param>
        /// <param name="outlookNote"></param>
        /// <param name="googleNote"></param>
        public static void SetOutlookGoogleNoteId(Syncronizer sync, Outlook.NoteItem outlookNote, Document googleNote)
        {
            if (googleNote.DocumentEntry.Id.Uri == null)
                throw new NullReferenceException("GoogleNote must have a valid Id");

            //check if outlook note aready has google id property.
            Outlook.ItemProperties userProperties = outlookNote.ItemProperties;
            try
            {
                Outlook.ItemProperty prop = userProperties[sync.OutlookPropertyNameId];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, true);
                try
                {
                    prop.Value = googleNote.DocumentEntry.Id.Uri.Content;
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }

            //save last google's updated date as property
            /*prop = outlookNote.UserProperties[OutlookPropertyNameUpdated];
            if (prop == null)
                prop = outlookNote.UserProperties.Add(OutlookPropertyNameUpdated, Outlook.OlUserPropertyType.olDateTime, null, null);
            prop.Value = googleNote.Updated;*/

            //Also set the OutlookLastSync date when setting a match between Outlook and Google to assure the lastSync updated when Outlook note is saved afterwards
            SetOutlookLastSync(sync, outlookNote);
        }

        public static void SetOutlookLastSync(Syncronizer sync, Outlook.NoteItem outlookNote)
        {
            //save sync datetime
            Outlook.ItemProperties userProperties = outlookNote.ItemProperties;
            try
            {
                Outlook.ItemProperty prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, true);
                try
                {
                    prop.Value = DateTime.Now;
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static DateTime? GetOutlookLastSync(Syncronizer sync, Outlook.NoteItem outlookNote)
        {
            DateTime? result = null;
            Outlook.ItemProperties userProperties = outlookNote.ItemProperties;
            try
            {
                Outlook.ItemProperty prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop != null)
                {
                    try
                    {
                        result = (DateTime)prop.Value;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(prop);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
            return result;
        }
        public static string GetOutlookGoogleNoteId(Syncronizer sync, Outlook.NoteItem outlookNote)
        {
            string id = null;
            Outlook.ItemProperties userProperties = outlookNote.ItemProperties;
            try
            {
                Outlook.ItemProperty idProp = userProperties[sync.OutlookPropertyNameId];
                if (idProp != null)
                {
                    try
                    {
                        id = (string)idProp.Value;
                        if (id == null)
                            throw new Exception();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(idProp);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
            return id;
        }
        public static void ResetOutlookGoogleNoteId(Syncronizer sync, Outlook.NoteItem outlookNote)
        {
            Outlook.ItemProperties userProperties = outlookNote.ItemProperties;
            try
            {
                Outlook.ItemProperty idProp = userProperties[sync.OutlookPropertyNameId];
                try
                {
                    Outlook.ItemProperty lastSyncProp = userProperties[sync.OutlookPropertyNameSynced];
                    try
                    {
                        if (idProp == null && lastSyncProp == null)
                            return;

                        List<int> indexesToBeRemoved = new List<int>();
                        IEnumerator en = userProperties.GetEnumerator();
                        en.Reset();
                        int index = 0; // 0 based collection            
                        while (en.MoveNext())
                        {
                            Outlook.ItemProperty userProperty = en.Current as Outlook.ItemProperty;
                            if (userProperty == idProp || userProperty == lastSyncProp)
                            {
                                indexesToBeRemoved.Add(index);
                                //outlookNote.UserProperties.Remove(index);
                                //Don't return to remove both properties, googleId and lastSynced
                                //return;
                            }
                            index++;
                            Marshal.ReleaseComObject(userProperty);
                        }

                        for (int i = indexesToBeRemoved.Count - 1; i >= 0; i--)
                            userProperties.Remove(indexesToBeRemoved[i]);
                        //throw new Exception("Did not find prop.");
                    }
                    finally
                    {
                        if (lastSyncProp != null)
                            Marshal.ReleaseComObject(lastSyncProp);
                    }
                }
                finally
                {
                    if (idProp != null)
                        Marshal.ReleaseComObject(idProp);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static string GetFileName(string Id, string syncProfile)
        {
            string fileName = "Note_" + Id + ".txt";


            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');                
            }

            //Only for backward compliance with version before 3.5.9 (before syncProfile can be changed)
            CopyNoteFiles(syncProfile);

            fileName = Logger.Folder + (string.IsNullOrEmpty(syncProfile)?string.Empty:"\\" + syncProfile) + "\\" + fileName;
            return fileName;
        }

        private static void CopyNoteFiles(string syncProfile)
        {
            //Only for backward compliance with version before 3.5.9 (before syncProfile can be changed)
            //Create ProfileSync subfolder and copy all files to there

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                syncProfile = syncProfile.Replace(c, '_');
            }

            if (!string.IsNullOrEmpty(syncProfile) && !Directory.Exists(Logger.Folder + "\\" + syncProfile))
            {
                Directory.CreateDirectory(Logger.Folder + "\\" + syncProfile);

                string[] files = Directory.GetFiles(Logger.Folder, @"Note_*.txt");

                foreach (string file in files)
                    File.Move(file, file.Replace(Logger.Folder, Logger.Folder + "\\" + syncProfile + "\\"));

            }
        }

        public static string GetBody(Syncronizer sync, Document entry)
        {
            string body = null;
            System.IO.StreamReader reader = null;
            try
            {
                reader = new System.IO.StreamReader(sync.DocumentsRequest.Download(entry, Document.DownloadType.txt));
                body = reader.ReadToEnd();
            }
            finally
            {
                if (reader != null)
                    reader.Close();
            }
            return body;
        }

        public static bool NoteFileExists(string Id, string syncProfile)
        {
            if (System.IO.File.Exists(GetFileName(Id, syncProfile)))
                return true;

            return false;
        }

        public static string CreateNoteFile(string Id, string body, string syncProfile)
        {
            string fileName = NotePropertiesUtils.GetFileName(Id, syncProfile);

            StreamWriter writer = null;
            try
            {
                FileStream filestream = new FileStream(fileName, FileMode.OpenOrCreate);
                writer = new StreamWriter(filestream);
                writer.Write(body);
            }
            finally
            {
                if (writer != null)
                    writer.Close();
            }

            return fileName;
        }

        public static void DeleteNoteFiles(string syncProfile)        
        {
            //Only for backward compliance with version before 3.5.9 (before syncProfile can be changed)
            CopyNoteFiles(syncProfile);

            string[] files = Directory.GetFiles(Logger.Folder + "\\" + syncProfile, @"Note_*.txt");

            foreach (string file in files)
                File.Delete(file);
        }

                
    }
}
