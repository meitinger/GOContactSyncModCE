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

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class GoogleAPITests
    {
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

            foreach (Contact entry in feed.Entries)
            {
                if (entry.PrimaryEmail != null && entry.PrimaryEmail.Address == "johndoe@example.com")
                {
                    service.Delete(entry);
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

            Assert.IsNotNull(createdEntry.ContactEntry.Id.Uri);

            //delete test contacts
            service.Delete(createdEntry);
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


            #region Delete previously created test note.
            DocumentQuery query = new DocumentQuery(service.BaseUri);
            query.NumberToRetrieve = 500;

            Feed<Document> feed = service.Get<Document>(query);

            foreach (Document entry in feed.Entries)
            {
                if (entry.Title == "AN_OUTLOOK_TEST_NOTE")
                {
                    //service.Delete(entry);
                    service.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + entry.ResourceId), entry.ETag); 
                    //break;
                }
            }
            #endregion


            Document newEntry = new Document();
            newEntry.Type = Document.DocumentType.Document;
            newEntry.Title = "AN_OUTLOOK_TEST_NOTE";

            string file = NotePropertiesUtils.CreateNoteFile("AN_OUTLOOK_TEST_NOTE", "This is just a test note to test GoContactSyncMod", null);
            newEntry.MediaSource = new MediaFileSource(file, MediaFileSource.GetContentTypeForFileName(file));

            #region normal flow, currently not working because of Note content
            Uri feedUri = new Uri(service.BaseUri);

            Document createdEntry = service.Insert(feedUri, newEntry);

            Assert.IsNotNull(createdEntry.DocumentEntry.Id.Uri);

            //delete test note            
            //ToDo: Doesn'T work always, frequently throwing 401, Precondition failed, maybe Google API bug
            //service.Delete(createdEntry);
            //Todo: Workaround to load document again
            feed = service.Get<Document>(query);

            foreach (Document entry in feed.Entries)
            {
                if (entry.Title == "AN_OUTLOOK_TEST_NOTE")
                {
                    service.Delete(entry);
                    break;
                }
            }            

            #endregion

            #region workaround flow to use UploadDocument
            Google.GData.Documents.DocumentEntry createdEntry2 = service.Service.UploadDocument(file, newEntry.Title);
            
            Assert.IsNotNull(createdEntry2.Id.Uri);

            //service.Service.Delete(createdEntry2);
            service.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + createdEntry2.ResourceId), createdEntry2.Etag); 
            #endregion

            System.IO.File.Delete(file);
        }

        internal static void LoadSettings(out string gmailUsername, out string gmailPassword, out string syncProfile, out string syncContactsFolder, out string syncNotesFolder)
        {
            Microsoft.Win32.RegistryKey regKeyAppRoot = LoadSettings(out gmailUsername, out gmailPassword, out syncProfile);

            syncContactsFolder = "";
            syncNotesFolder = "";
            if (regKeyAppRoot.GetValue("SyncContactsFolder") != null)
                syncContactsFolder = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
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
