using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Web;
using System.Net;
using System.IO;
using System.Drawing;
using System.Configuration;
using Google.Documents;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncNotesTests
    {
        Syncronizer sync;

        static Logger.LogUpdatedHandler _logUpdateHandler = null;

        //Constants for test Note
        const string name = "AN_OUTLOOK_TEST_NOTE";
        const string body = "This is just a test note to test GoContactSyncMod";
        const string groupName = "A TEST GROUP";

        [TestFixtureSetUp]
        public void Init() 
        {            
            //string timestamp = DateTime.Now.Ticks.ToString();            
            if (_logUpdateHandler == null)
            {
                _logUpdateHandler = new Logger.LogUpdatedHandler(Logger_LogUpdated);
                Logger.LogUpdated += _logUpdateHandler;
            }

            string gmailUsername;
            string gmailPassword;
            string syncProfile;
            string syncContactsFolder;
            string syncNotesFolder;

            GoogleAPITests.LoadSettings(out gmailUsername, out gmailPassword, out syncProfile, out syncContactsFolder, out syncNotesFolder);

            sync = new Syncronizer();
            sync.SyncContacts = false;
            sync.SyncNotes = true;
            sync.SyncProfile = syncProfile;
            Syncronizer.SyncNotesFolder = syncNotesFolder;           

            sync.LoginToGoogle(gmailUsername, gmailPassword);
            sync.LoginToOutlook();

            

        }        

        [SetUp]
        public void SetUp()
        {
            // delete previously failed test Notes
            DeleteTestNotes();
                      
        }

        private void DeleteTestNotes()
        {
            sync.LoadNotes();

            //Outlook.NoteItem outlookNote = sync.OutlookNotes.Find("[Body] = '" + body + "'") as Outlook.NoteItem;
            foreach (Outlook.NoteItem outlookNote in sync.OutlookNotes)
            {
                if (outlookNote != null &&
                    outlookNote.Body != null && outlookNote.Body == body)
                    DeleteTestNote(outlookNote);
            }

            foreach (Document googleNote in sync.GoogleNotes)
            {
                string noteBody = NotePropertiesUtils.GetBody(sync, googleNote);
                if (googleNote != null &&
                    noteBody != null && noteBody == body)
                {
                    DeleteTestNote(googleNote);
                }
            }

            //Delete empty Google note folders
            sync.CleanUpGoogleCategories();
        }

        void Logger_LogUpdated(string message)
        {
            Console.WriteLine(message);
        }

        [TestFixtureTearDown]        
        public void TearDown()
        {
            sync.LogoffOutlook();
            sync.LogoffGoogle();
        }

        
        [Test]
        public void TestSync()
        {
            // create new note to sync
            Outlook.NoteItem outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            //outlookNote.Subject = name;          
            outlookNote.Body = body;
           
            outlookNote.Save();

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;     

            Document googleNote = new Document();
            googleNote.Type = Document.DocumentType.Document;
            sync.UpdateNote(outlookNote, googleNote);
            NoteMatch match = new NoteMatch(outlookNote, googleNote);

            //save Note to google.
            sync.SaveGoogleNote(match);
            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 10; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds

            googleNote = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;   
            //load the same Note from google.
            sync.MatchNotes();
            match = FindMatch(match.GoogleNote);
            //NotesMatcher.SyncNote(match, sync);

            Outlook.NoteItem recreatedOutlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            sync.UpdateNote(match.GoogleNote, recreatedOutlookNote);

            // match recreatedOutlookNote with outlookNote
            //Assert.AreEqual(outlookNote.Subject, recreatedOutlookNote.Subject);           
            Assert.AreEqual(outlookNote.Body, recreatedOutlookNote.Body);
           
            DeleteTestNotes(match);
        }

        [Test]
        public void TestSyncDeletedOulook()
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new Note to sync
            Outlook.NoteItem outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            outlookNote.Body = body;            
            outlookNote.Save();

            Document googleNote = new Document();
            googleNote.Type = Document.DocumentType.Document;
            sync.UpdateNote(outlookNote, googleNote);
            NoteMatch match = new NoteMatch(outlookNote, googleNote);

            //save Notes
            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 10; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds

            // delete outlook Note
            outlookNote.Delete();

            Thread.Sleep(10000);

            // sync
            sync.MatchNotes();
            NotesMatcher.SyncNotes(sync);
            // find match
            match = FindMatch(match.GoogleNote);            

            // delete
            sync.SaveNote(match);

            // sync
            sync.MatchNotes();
            NotesMatcher.SyncNotes(sync);

            // check if google Note still exists
            match = FindMatch(match.GoogleNote);
            
            Assert.IsNull(match);
        }

        [Test]
        public void TestSyncDeletedGoogle()
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;            
            sync.SyncDelete = true;

            // create new Note to sync
            Outlook.NoteItem outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            outlookNote.Body = body;            
            outlookNote.Save();

            Document googleNote = new Document();
            googleNote.Type = Document.DocumentType.Document;
            sync.UpdateNote(outlookNote, googleNote);
            NoteMatch match = new NoteMatch(outlookNote, googleNote);

            //save Notes
            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 100; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds

            Document deletedNote = sync.LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
            Assert.IsNotNull(deletedNote);
            AtomId deletedNoteAtomId = deletedNote.DocumentEntry.Id;
            string deletedNoteId = deletedNote.Id;

            Assert.IsTrue(File.Exists(NotePropertiesUtils.GetFileName(deletedNoteId, sync.SyncProfile)));

            // delete google Note
            sync.DocumentsRequest.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + deletedNote.ResourceId), deletedNote.ETag); 

            // sync
            sync.MatchNotes();
            match = FindMatch(outlookNote);
            NotesMatcher.SyncNote(match, sync);

            string id = outlookNote.EntryID;

            // delete
            sync.SaveNote(match);

            // sync
            sync.MatchNotes();
            NotesMatcher.SyncNotes(sync);
            match = FindMatch(id);            

            // check if outlook Note still exists
            Assert.IsNull(match);

            deletedNote = sync.LoadGoogleNotes(deletedNoteAtomId);
            Assert.IsNull(deletedNote);

            Assert.IsFalse(File.Exists(NotePropertiesUtils.GetFileName(deletedNoteId, sync.SyncProfile)));
            Assert.IsFalse(File.Exists(NotePropertiesUtils.GetFileName(id, sync.SyncProfile)));

            DeleteTestNotes(match);
                      
        }

        [Test]
        public void TestSyncGroups()
        {
            Outlook.NoteItem outlookNote;
            NoteMatch match;

            CreateOutlookNoteAndSyncToGoogle(out outlookNote, out match);

            // delete outlook note
            outlookNote.Delete();
            outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            sync.UpdateNote(match.GoogleNote, outlookNote);
            match = new NoteMatch(outlookNote, match.GoogleNote);
            outlookNote.Save();

            sync.SyncOption = SyncOption.MergeGoogleWins;

            //sync and save contact to outlook
            sync.UpdateNote(match.GoogleNote, outlookNote);
            sync.SaveNote(match);

            //load the same contact from outlook
            sync.MatchNotes();
            match = FindMatch(outlookNote);

            Assert.AreEqual(groupName, outlookNote.Categories);

            DeleteTestNotes(match);

            //Delete empty Google note folders
            sync.CleanUpGoogleCategories();
        }

        [Test]
        public void TestSyncDeletedGroup()
        {
            Outlook.NoteItem outlookNote;
            NoteMatch match;
            
            CreateOutlookNoteAndSyncToGoogle(out outlookNote, out match);

            // remove category from Outlook Note
            outlookNote.Categories = null;

            sync.UpdateNote(match.OutlookNote, match.GoogleNote);

            //save Note to Google
            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 100; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds                      

            //load the same note from google.
            sync.MatchNotes();
            match = FindMatch(outlookNote);

            // google contact should now have no group anymore                     
            Assert.AreEqual(1, match.GoogleNote.ParentFolders.Count);   //Only Notes Folder remains
            Assert.AreEqual(match.GoogleNote.ParentFolders[0], sync.googleNotesFolder.Self);

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;

            outlookNote.Categories = groupName;

            sync.UpdateNote(match.GoogleNote, match.OutlookNote);

            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 100; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds

            //Sync Groups first
            sync.MatchNotes();

            //sync and save contact to outlook.
            match = FindMatch(outlookNote);
            sync.UpdateNote(match.GoogleNote, outlookNote);
            sync.SaveNote(match);

            // google and outlook should now have no category            
            Assert.AreEqual(1, match.GoogleNote.ParentFolders.Count);   //Only Notes Folder remains
            Assert.AreEqual(match.GoogleNote.ParentFolders[0], sync.googleNotesFolder.Self);
            Assert.AreEqual(null, outlookNote.Categories);
           

            DeleteTestNotes(match);

            //Delete empty Google note folders
            sync.CleanUpGoogleCategories();
        }

        private void CreateOutlookNoteAndSyncToGoogle(out Outlook.NoteItem outlookNote, out NoteMatch match)
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            // create new note to sync
            outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            outlookNote.Body = body;
            outlookNote.Categories = groupName;
            outlookNote.Save();

            //Outlook note should now have a group
            Assert.AreEqual(groupName, outlookNote.Categories);

            Document googleNote = new Document();
            googleNote.Type = Document.DocumentType.Document;
            sync.UpdateNote(outlookNote, googleNote);
            match = new NoteMatch(outlookNote, googleNote);

            //save Notes
            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 100; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds                      

            //load the same note from google.
            sync.MatchNotes();
            match = FindMatch(outlookNote);

            // google contact should now have the same group            
            Assert.IsNotNull(match.GoogleNote.ParentFolders);
            //Assert.Greater(match.GoogleNote.ParentFolders.Count, 0); 
            Assert.AreEqual(2, match.GoogleNote.ParentFolders.Count);   //2 because Notes folder and the category folder 

            Document categoryFolder = null;
            foreach (string uri in match.GoogleNote.ParentFolders)
            {
                Document folder = sync.GetGoogleFolder(null, null, uri);
                if (folder.Title == groupName)
                {
                    categoryFolder = folder;
                    break;
                }
            }
            Assert.IsNotNull(categoryFolder);
        }
        

       
        [Test]
        public void TestResetMatches()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new Note to sync
            Outlook.NoteItem outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            outlookNote.Body = body;           
            outlookNote.Save();

            Document googleNote = new Document();
            googleNote.Type = Document.DocumentType.Document;
            sync.UpdateNote(outlookNote, googleNote);
            NoteMatch match = new NoteMatch(outlookNote, googleNote);

            //save Note to google.
            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 10; i++ )
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds

            //load the same Note from google.
            sync.MatchNotes();
            match = FindMatch(outlookNote);
            NotesMatcher.SyncNote(match, sync);

            // delete outlook Note
            outlookNote.Delete();
            match.OutlookNote = null;

            //load the same Note from google
            sync.MatchNotes();
            match = FindMatch(match.GoogleNote);
            NotesMatcher.SyncNote(match, sync);

            Assert.IsNull(match.OutlookNote);

            // reset matches
            System.IO.File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id, sync.SyncProfile));
            //Not, because NULL: sync.ResetMatch(match.OutlookNote.GetOriginalItemFromOutlook(sync));
            
            // load same Note match
            sync.MatchNotes();
            match = FindMatch(match.GoogleNote);
            NotesMatcher.SyncNote(match, sync);

            // google Note should still be present and OutlookNote should be filled
            Assert.IsNotNull(match.GoogleNote);
            Assert.IsNotNull(match.OutlookNote);

            DeleteTestNotes();    

            // create new Note to sync
            outlookNote = Syncronizer.CreateOutlookNoteItem(Syncronizer.SyncNotesFolder);
            outlookNote.Body = body;          
            outlookNote.Save();

            // same test for delete google Note...
            googleNote = new Document();
            googleNote.Type = Document.DocumentType.Document;
            sync.UpdateNote(outlookNote, googleNote);
            match = new NoteMatch(outlookNote, googleNote);

            //save Note to google.
            sync.SaveNote(match);

            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < 10; i++)
                Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds

            //load the same Note from google.
            sync.MatchNotes();
            match = FindMatch(outlookNote);
            NotesMatcher.SyncNote(match, sync);

            // delete google Note           
            //sync.DocumentsRequest.Delete(match.GoogleNote);   
            DeleteTestNote(match.GoogleNote);
            match.GoogleNote = null;

            //load the same Note from google.
            sync.MatchNotes();
            match = FindMatch(outlookNote);
            NotesMatcher.SyncNote(match, sync);

            Assert.IsNull(match.GoogleNote);

            // reset matches
            //Not, because null: sync.ResetMatch(match.GoogleNote);
            sync.ResetMatch(match.OutlookNote);

            // load same Note match
            sync.MatchNotes();
            match = FindMatch(outlookNote);
            NotesMatcher.SyncNote(match, sync);

            // Outlook Note should still be present and GoogleNote should be filled
            Assert.IsNotNull(match.OutlookNote);
            Assert.IsNotNull(match.GoogleNote);

            System.IO.File.Delete(NotePropertiesUtils.GetFileName(outlookNote.EntryID, sync.SyncProfile));
            outlookNote.Delete();            
        }

        private void DeleteTestNotes(NoteMatch match)
        {
            if (match != null)
            {
                DeleteTestNote(match.GoogleNote);
                DeleteTestNote(match.OutlookNote);
            }
        }

        private void DeleteTestNote(Outlook.NoteItem outlookNote)
        {
            if (outlookNote != null)
            {
                try
                {
                    outlookNote.Delete();
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNote);
                    outlookNote = null;
                }
                
            }
        }


        private void DeleteTestNote(Document googleNote)
        {
            if (googleNote != null)
            {
                sync.DocumentsRequest.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + googleNote.ResourceId), googleNote.ETag);
                //sync.DocumentsRequest.Delete(googleNote);

                ////ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder, therefore the following workaround
                //Document deletedNote = sync.LoadGoogleNotes(googleNote.DocumentEntry.Id);
                //if (deletedNote != null)
                //    sync.DocumentsRequest.Delete(deletedNote);

                try
                {//Delete also the according temporary NoteFile
                    File.Delete(NotePropertiesUtils.GetFileName(googleNote.Id, sync.SyncProfile));
                }
                catch (System.Exception)
                { }
            }
        }
        
        internal NoteMatch FindMatch(Outlook.NoteItem outlookNote)
        {
            return FindMatch(outlookNote.EntryID);
        }

        internal NoteMatch FindMatch(string outlookNoteId)
        {
            foreach (NoteMatch match in sync.Notes)
            {
                if (match.OutlookNote.EntryID == outlookNoteId)
                    return match;
            }
            return null;
        }

        internal NoteMatch FindMatch(Document googleNote)
        {
            if (googleNote != null)
            {
                foreach (NoteMatch match in sync.Notes)
                {
                    if (match.GoogleNote != null && match.GoogleNote.DocumentEntry.Id == googleNote.DocumentEntry.Id)
                        return match;
                }
            }
            return null;
        }

    }
}
