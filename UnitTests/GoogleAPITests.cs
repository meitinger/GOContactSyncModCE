using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using System.Configuration;
using Google.Contacts;
using Google.Documents;
using Google.GData.Client.ResumableUpload;
using Google.GData.Documents;
using System.Collections;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class GoogleAPITests
    {
        ClientLoginAuthenticator _authenticator;
        static Logger.LogUpdatedHandler _logUpdateHandler = null;
        void Logger_LogUpdated(string message)
        {
            Console.WriteLine(message);
        }

        [TestFixtureSetUp]
        public void Init()
        {
            //string timestamp = DateTime.Now.Ticks.ToString();            
            if (_logUpdateHandler == null)
            {
                _logUpdateHandler = new Logger.LogUpdatedHandler(Logger_LogUpdated);
                Logger.LogUpdated += _logUpdateHandler;
            }
        }    

        [Test]
        public void CreateNewContact()
        {
            string gmailUsername;
            string gmailPassword;
            string syncProfile;
            GoogleAPITests.LoadSettings(out gmailUsername, out gmailPassword, out syncProfile);

            RequestSettings rs = new RequestSettings("GoogleContactSyncMod", gmailUsername, gmailPassword);
            ContactsRequest service = new ContactsRequest(rs);


            #region Delete previously created test contact.
            ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 500;

            Feed<Contact> feed = service.Get<Contact>(query);

            Logger.Log("Loaded Google contacts", EventType.Information);

            foreach (Contact entry in feed.Entries)
            {
                if (entry.PrimaryEmail != null && entry.PrimaryEmail.Address == "johndoe@example.com")
                {
                    service.Delete(entry);
                    Logger.Log("Deleted Google contact", EventType.Information);
                    //break;
                }
            }
            #endregion

            Contact newEntry = new Contact();
            newEntry.Title = "John Doe";

            EMail primaryEmail = new EMail("johndoe@example.com");
            primaryEmail.Primary = true;
            primaryEmail.Rel = ContactsRelationships.IsWork;
            newEntry.Emails.Add(primaryEmail);

            PhoneNumber phoneNumber = new PhoneNumber("555-555-5551");
            phoneNumber.Primary = true;
            phoneNumber.Rel = ContactsRelationships.IsMobile;
            newEntry.Phonenumbers.Add(phoneNumber);

            StructuredPostalAddress postalAddress = new StructuredPostalAddress();
            postalAddress.Street = "123 somewhere lane";
            postalAddress.Primary = true;
            postalAddress.Rel = ContactsRelationships.IsHome;
            newEntry.PostalAddresses.Add(postalAddress);

            newEntry.Content = "Who is this guy?";

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            Contact createdEntry = service.Insert(feedUri, newEntry);

            Logger.Log("Created Google contact", EventType.Information);

            Assert.IsNotNull(createdEntry.ContactEntry.Id.Uri);

            Contact updatedEntry = service.Update(createdEntry);

            Logger.Log("Updated Google contact", EventType.Information);

            //delete test contacts
            service.Delete(createdEntry);

            Logger.Log("Deleted Google contact", EventType.Information);
        }


        [Test]
        public void CreateNewNote()
        {
            string gmailUsername;
            string gmailPassword;
            string syncProfile;
            GoogleAPITests.LoadSettings(out gmailUsername, out gmailPassword, out syncProfile);

            RequestSettings rs = new RequestSettings("GoogleContactSyncMod", gmailUsername, gmailPassword);
            DocumentsRequest service = new DocumentsRequest(rs);
            //Instantiate an Authenticator object according to your authentication, to use ResumableUploader
            _authenticator = new ClientLoginAuthenticator("GCSM Unit Tests", service.Service.ServiceIdentifier, gmailUsername, gmailPassword);

            //Delete previously created test note.            
            DeleteTestNote(service);

            Document newEntry = new Document();
            newEntry.Type = Document.DocumentType.Document;
            newEntry.Title = "AN_OUTLOOK_TEST_NOTE";

            string file = NotePropertiesUtils.CreateNoteFile("AN_OUTLOOK_TEST_NOTE", "This is just a test note to test GoContactSyncMod", null);
            newEntry.MediaSource = new MediaFileSource(file, MediaFileSource.GetContentTypeForFileName(file));

            #region normal flow, only working to create documents without content (metadata only), e.g. for Notes folder

            Document createdEntry = Syncronizer.SaveGoogleNote(null, newEntry, service);

            Assert.IsNotNull(createdEntry.DocumentEntry.Id.Uri);

            Logger.Log("Created Google note", EventType.Information);

            //delete test note            
            DeleteTestNote(service);
            #endregion

            #region workaround flow to use UploadDocument, not needed anymore because of new approach to use ResumableUploader
            //Google.GData.Documents.DocumentEntry createdEntry2 = service.Service.UploadDocument(file, newEntry.Title);
            
            //Assert.IsNotNull(createdEntry2.Id.Uri);

            ////delete test note            
            //DeleteTestNote(service);
            #endregion

            #region New approach how to update an existing document: https://developers.google.com/google-apps/documents-list/#updatingchanging_documents_and_files            
            //Instantiate the ResumableUploader component.      
            ResumableUploader uploader = new ResumableUploader();            
            uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteCreated);
            Syncronizer.CreateGoogleNote(newEntry, file, service, uploader, _authenticator);            
            #endregion            

            //Wait 5 seconds to give the testcase the chance to finish the Async events
            System.Threading.Thread.Sleep(5000);

            DeleteTestNote(service);
        }

        private void OnGoogleNoteCreated(object sender, AsyncOperationCompletedEventArgs e)
        {
            DocumentEntry entry = e.Entry as DocumentEntry;

            Assert.IsNotNull(entry);

            Logger.Log("Created Google note", EventType.Information);

            //Now update the same entry
            //Instantiate the ResumableUploader component.      
            ResumableUploader uploader = new ResumableUploader();
            uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteUpdated);
            uploader.UpdateAsync(_authenticator, entry, e.UserState);


        }

        private void OnGoogleNoteUpdated(object sender, AsyncOperationCompletedEventArgs e)
        {
            DocumentEntry entry = e.Entry as DocumentEntry;
            
            Assert.IsNotNull(entry);

            Logger.Log("Updated Google note", EventType.Information);
            
            System.IO.File.Delete(e.UserState as string);
        }

        private static void DeleteTestNote(DocumentsRequest service)
        {
            //ToDo: Doesn'T work always, frequently throwing 401, Precondition failed, maybe Google API bug
            //service.Delete(createdEntry);

            //Todo: Workaround to load document again
            DocumentQuery query = new DocumentQuery(service.BaseUri);
            query.NumberToRetrieve = 500;

            Feed<Document> feed = service.Get<Document>(query);

            Logger.Log("Loaded Google notes", EventType.Information);

            foreach (Document entry in feed.Entries)
            {
                if (entry.Title == "AN_OUTLOOK_TEST_NOTE")
                {
                    //service.Delete(entry);
                    service.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + entry.ResourceId), entry.ETag);
                    Logger.Log("Deleted Google note", EventType.Information);
                    //break;
                }
            }
        }

        internal static void LoadSettings(out string gmailUsername, out string gmailPassword, out string syncProfile, out string syncContactsFolder, out string syncNotesFolder)
        {
            Microsoft.Win32.RegistryKey regKeyAppRoot = LoadSettings(out gmailUsername, out gmailPassword, out syncProfile);

            syncContactsFolder = "";
            syncNotesFolder = "";

            //First, check if there is a folder called GCSMTestContacts and GCSMTestNotes available, if yes, use them
            ArrayList outlookContactFolders = new ArrayList();
            ArrayList outlookNoteFolders = new ArrayList();
            Microsoft.Office.Interop.Outlook.Folders folders = Syncronizer.OutlookNameSpace.Folders;
            foreach (Microsoft.Office.Interop.Outlook.Folder folder in folders)
            {
                try
                {
                    SettingsForm.GetOutlookMAPIFolders(outlookContactFolders, outlookNoteFolders, folder);
                }
                catch (Exception e)
                {
                    Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                }
            }

            foreach (OutlookFolder folder in outlookContactFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestContacts".ToUpper()))
                {
                    Logger.Log("Uses Test folder: " + folder.DisplayName, EventType.Information);
                    syncContactsFolder = folder.FolderID;
                    break;
                }
            }
            foreach (OutlookFolder folder in outlookNoteFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestNotes".ToUpper()))
                {
                    Logger.Log("Uses Test folder: " + folder.DisplayName, EventType.Information);
                    syncNotesFolder = folder.FolderID;
                    break;
                }
            }

            if (string.IsNullOrEmpty(syncContactsFolder))
                if (regKeyAppRoot.GetValue("SyncContactsFolder") != null)
                    syncContactsFolder = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
            if (string.IsNullOrEmpty(syncNotesFolder))
                if (regKeyAppRoot.GetValue("SyncNotesFolder") != null)
                    syncNotesFolder = regKeyAppRoot.GetValue("SyncNotesFolder") as string;           
        }

        private static Microsoft.Win32.RegistryKey LoadSettings(out string gmailUsername, out string gmailPassword, out string syncProfile)
        {
            //sync.LoginToGoogle(ConfigurationManager.AppSettings["Gmail.Username"],  ConfigurationManager.AppSettings["Gmail.Password"]);
            //ToDo: Reading the username and config from the App.Config file doesn't work. If it works, consider special characters like & = &amp; in the XML file
            //ToDo: Maybe add a common Test account to the App.config and read it from there, encrypt the password
            //For now, read the userName and Password from the Registry (same settings as for GoogleContactsSync Application
            gmailUsername = "";
            gmailPassword = "";

            const string appRootKey = @"Software\Webgear\GOContactSync";
            Microsoft.Win32.RegistryKey regKeyAppRoot = regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRootKey);
            syncProfile = "Default Profile";
            if (regKeyAppRoot.GetValue("SyncProfile") != null)
                syncProfile = regKeyAppRoot.GetValue("SyncProfile") as string;

            regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRootKey + (syncProfile != null ? ('\\' + syncProfile) : ""));

            if (regKeyAppRoot.GetValue("Username") != null)
            {
                gmailUsername = regKeyAppRoot.GetValue("Username") as string;
                if (regKeyAppRoot.GetValue("Password") != null)
                    gmailPassword = Encryption.DecryptPassword(gmailUsername, regKeyAppRoot.GetValue("Password") as string);
            }

            return regKeyAppRoot;
        }
    }

}
