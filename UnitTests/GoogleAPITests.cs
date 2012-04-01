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
            string gmailUsername = "";
            string gmailPassword = "";

            Microsoft.Win32.RegistryKey regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            if (regKeyAppRoot.GetValue("Username") != null)
            {
                gmailUsername = regKeyAppRoot.GetValue("Username") as string;
                if (regKeyAppRoot.GetValue("Password") != null)
                    gmailPassword = Encryption.DecryptPassword(gmailUsername, regKeyAppRoot.GetValue("Password") as string);
            }

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
                    break;
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
            string gmailUsername = "";
            string gmailPassword = "";

            Microsoft.Win32.RegistryKey regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            if (regKeyAppRoot.GetValue("Username") != null)
            {
                gmailUsername = regKeyAppRoot.GetValue("Username") as string;
                if (regKeyAppRoot.GetValue("Password") != null)
                    gmailPassword = Encryption.DecryptPassword(gmailUsername, regKeyAppRoot.GetValue("Password") as string);
            }

            RequestSettings rs = new RequestSettings("GoogleContactSyncMod", gmailUsername, gmailPassword);
            DocumentsRequest service = new DocumentsRequest(rs);


            #region Delete previously created test note.
            DocumentQuery query = new DocumentQuery(service.BaseUri);
            query.NumberToRetrieve = 500;

            Feed<Document> feed = service.Get<Document>(query);

            foreach (Document entry in feed.Entries)
            {
                if (entry.Title == "OutlookTestNote")
                {
                    service.Delete(entry);
                    break;
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

            service.Service.Delete(createdEntry2);
            #endregion

            System.IO.File.Delete(file);
        }
    }

}
