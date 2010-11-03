using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Google.GData.Contacts;
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

namespace WebGear.GoogleContactsSync.UnitTests
{
    [TestFixture]
    public class SyncTests
    {
        Syncronizer sync;

        [SetUp]
        public void SetUp()
        {
            sync = new Syncronizer();
            sync.SyncProfile = "test profile";
            sync.LoginToGoogle(ConfigurationManager.AppSettings["Gmail.Username"], ConfigurationManager.AppSettings["Gmail.Password"]);
            sync.LoginToOutlook();
            sync.Load();
        }

        [TearDown]
        public void TearDown()
        {
            sync.LogoffOutlook();
        }

        [Test]
        public void TestSync()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            Outlook.ContactItem recreatedOutlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            ContactSync.MergeContacts(match.GoogleContact, recreatedOutlookContact);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(outlookContact.FullName, recreatedOutlookContact.FullName);
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);
            Assert.AreEqual(outlookContact.Email1Address, recreatedOutlookContact.Email1Address);
            Assert.AreEqual(outlookContact.Email2Address, recreatedOutlookContact.Email2Address);
            Assert.AreEqual(outlookContact.Email3Address, recreatedOutlookContact.Email3Address);
            Assert.AreEqual(outlookContact.PrimaryTelephoneNumber, recreatedOutlookContact.PrimaryTelephoneNumber);
            Assert.AreEqual(outlookContact.HomeAddress, recreatedOutlookContact.HomeAddress);

            //delete test contacts
            match.Delete();
        }

        [Test]
        public void TestExtendedProps()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            sync.SaveGoogleContact(match);

            // read contact from google
            googleContact = null;
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            foreach (ContactMatch m in sync.Contacts)
            {
                if (m.GoogleContact.Title.Text == name)
                {
                    googleContact = m.GoogleContact;
                    break;
                }
            }

            // get extended prop
            Assert.AreEqual(ContactPropertiesUtils.GetOutlookId(outlookContact), ContactPropertiesUtils.GetGoogleOutlookContactId(sync.SyncProfile, googleContact));

            //delete test contacts
            match.Delete();
        }

        [Test]
        public void TestSyncDeletedOulook()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contacts
            sync.SaveContact(match);

            // delete outlook contact
            outlookContact.Delete();

            // sync
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            // find match
            match = null;
            foreach (ContactMatch m in sync.Contacts)
            {
                if (m.GoogleContact.Title.Text == name)
                {
                    match = m;
                    break;
                }
            }

            // delete
            sync.SaveContact(match);

            // sync
            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // check if google contact still exists
            googleContact = null;
            foreach (ContactMatch m in sync.Contacts)
            {
                if (m.GoogleContact.Title.Text == name)
                {
                    googleContact = m.GoogleContact;
                    break;
                }
            }
            Assert.IsNull(googleContact);
        }

        [Test]
        public void TestSyncDeletedGoogle()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contacts
            sync.SaveContact(match);

            // delete google contact
            match.GoogleContact.Delete();

            // sync
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // delete
            sync.SaveContact(match);

            // sync
            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // check if outlook contact still exists
            match = sync.ContactByProperty(name, email);
            Assert.IsNull(match);
        }

        [Test]
        public void TestGooglePhoto()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            Image pic = Utilities.CropImageGoogleFormat(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));
            bool saved = Utilities.SaveGooglePhoto(sync, match.GoogleContact, pic);
            Assert.IsTrue(saved);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            Image image = Utilities.GetGooglePhoto(sync, match.GoogleContact);
            Assert.IsNotNull(image);

            //delete test contacts
            match.Delete();
        }

        [Test]
        public void TestOutlookPhoto()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            ContactMatch match = sync.ContactByProperty(name, email);

            bool saved = Utilities.SetOutlookPhoto(match.OutlookContact, AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            Assert.IsTrue(saved);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            Image image = Utilities.GetOutlookPhoto(match.OutlookContact);
            Assert.IsNotNull(image);

            match.Delete();
        }

        [Test]
        public void TestSyncPhoto()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            Utilities.SetOutlookPhoto(outlookContact, AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // google contact should now have a photo
            Assert.IsNotNull(Utilities.GetGooglePhoto(sync, match.GoogleContact));

            // delete outlook contact
            match.OutlookContact.Delete();
            outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            ContactSync.UpdateContact(match.GoogleContact, outlookContact);
            match = new ContactMatch(outlookContact, match.GoogleContact);
            match.OutlookContact.Save();

            //save contact to outlook.
            sync.SaveContact(match);

            // outlook contact should now have a photo
            Assert.IsNotNull(Utilities.GetOutlookPhoto(match.OutlookContact));

            //delete test contacts
            match.Delete();
        }

        [Test]
        public void TestSyncGroups()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string groupName = "A_TEST_GROUP";
            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Categories = groupName;
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // google contatc should now have the same group
            Assert.AreEqual(sync.GetGoogleGroupByName(groupName), Utilities.GetGoogleGroups(sync, match.GoogleContact)[0]);

            // delete outlook contact
            match.OutlookContact.Delete();
            outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            ContactSync.UpdateContact(match.GoogleContact, outlookContact);
            match = new ContactMatch(outlookContact, match.GoogleContact);
            match.OutlookContact.Save();

            sync.SyncOption = SyncOption.MergeGoogleWins;

            //save contact to outlook.
            sync.SaveContact(match);

            // outlook contact should now have the same group
            Assert.AreEqual(groupName, match.OutlookContact.Categories);

            //delete test contacts
            match.Delete();

            // delete test group
            GroupEntry group = sync.GetGoogleGroupByName(groupName);
            group.Delete();
        }

        [Test]
        public void TestSyncDeletedGoogleGroup()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string groupName = "A_TEST_GROUP";
            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Categories = groupName;
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // delete group from google
            Utilities.RemoveGoogleGroup(match.GoogleContact, sync.GetGoogleGroupByName(groupName));

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // google and outlook should now have no category
            Assert.AreEqual(0, Utilities.GetGoogleGroups(sync, match.GoogleContact).Count);
            Assert.AreEqual(null, match.OutlookContact.Categories);

            //delete test contacts
            match.Delete();

            // delete test group
            GroupEntry group = sync.GetGoogleGroupByName(groupName);
            group.Delete();
        }

        [Test]
        public void TestSyncDeletedOutlookGroup()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string groupName = "A_TEST_GROUP";
            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Categories = groupName;
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // delete group from outlook
            Utilities.RemoveOutlookGroup(match.OutlookContact, groupName);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // google and outlook should now have no category
            Assert.AreEqual(null, match.OutlookContact.Categories);
            Assert.AreEqual(0, Utilities.GetGoogleGroups(sync, match.GoogleContact).Count);

            //delete test contacts
            match.Delete();

            // delete test group
            GroupEntry group = sync.GetGoogleGroupByName(groupName);
            group.Delete();
        }

        [Test]
        public void TestResetMatches()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            string groupName = "A_TEST_GROUP";
            string timestamp = DateTime.Now.Ticks.ToString();
            string name = "AN OUTLOOK TEST CONTACT";
            string email = "email00@outlook.com";
            name = name.Replace(" ", "_");

            // delete previously failed test contacts
            DeleteExistingTestContacts(name, email);

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // create new contact to sync
            Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            ContactEntry googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // delete outlook contact
            match.OutlookContact.Delete();
            match.OutlookContact = null;

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // reset matches
            sync.ResetMatch(match);

            // load same contact match
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // google contact should still be present
            Assert.IsNotNull(match.GoogleContact);

            //delete test contacts
            match.Delete();

            // create new contact to sync
            outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            // same test for delete google contact...
            googleContact = new ContactEntry();
            ContactSync.UpdateContact(outlookContact, googleContact);
            match = new ContactMatch(outlookContact, googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // delete google contact
            match.GoogleContact.Delete();
            match.GoogleContact = null;

            //load the same contact from google.
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // reset matches
            sync.ResetMatch(match);

            // load same contact match
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);

            // google contact should still be present
            Assert.IsNotNull(match.OutlookContact);

            //delete test contacts
            match.Delete();
        }

        //[Test]
        //public void TestMultiProfileSync()
        //{
        //    sync.SyncOption = SyncOption.MergeOutlookWins;

        //    string timestamp = DateTime.Now.Ticks.ToString();
        //    string name = "AN OUTLOOK TEST CONTACT";
        //    string email = "email00@outlook.com";
        //    name = name.Replace(" ", "_");

        //    // delete previously failed test contacts
        //    DeleteExistingTestContacts(name, email);

        //    sync.Load();
        //    ContactsMatcher.SyncContacts(sync);

        //    // create new contact to sync
        //    Outlook.ContactItem outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
        //    outlookContact.FullName = name;
        //    outlookContact.FileAs = name;
        //    outlookContact.Email1Address = email;
        //    outlookContact.Email2Address = email.Replace("00", "01");
        //    outlookContact.Email3Address = email.Replace("00", "02");
        //    outlookContact.HomeAddress = "10 Parades";
        //    outlookContact.PrimaryTelephoneNumber = "123";
        //    outlookContact.Save();

        //    ContactEntry googleContact = new ContactEntry();
        //    ContactSync.UpdateContact(outlookContact, googleContact);
        //    ContactMatch match = new ContactMatch(outlookContact, googleContact);

        //    //save contacts
        //    sync.SaveContact(match);

        //    // delete outlook contact
        //    outlookContact.Delete();

        //    // sync with different profile
        //    sync.SyncProfile = "test profile 2";
        //    sync.Load();
        //    ContactsMatcher.SyncContacts(sync);
        //    // find match
        //    match = null;
        //    match = sync.ContactByProperty(name, email);
        //    sync.SaveContact(match);

        //    // there should now be a contact under the new profile
        //    match = sync.ContactByProperty(name, email);
        //    Assert.IsNotNull(match.OutlookContact);
            
        //    // now delete the original outlook contact
        //    sync.SyncProfile = "test profile";
        //    match.OutlookContact.Delete();


        //}

        [Ignore]
        public void TestMassSyncToGoogle()
        {
            // NEED TO DELETE CONTACTS MANUALY

            int c = 300;
            string[] names = new string[c];
            string[] emails = new string[c];
            Outlook.ContactItem outlookContact;
            ContactMatch match;

            for (int i = 0; i < c; i++)
            {
                names[i] = "TEST name" + i;
                emails[i] = "email" + i + "@domain.com";
            }

            // count existing google contacts
            int excount = sync.GoogleContacts.Count;

            DateTime start = DateTime.Now;
            Console.WriteLine("Started mass sync to google of " + c + " contacts");

            for (int i = 0; i < c; i++)
            {
                outlookContact = sync.OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
                outlookContact.FullName = names[i];
                outlookContact.FileAs = names[i];
                outlookContact.Email1Address = emails[i];
                outlookContact.Save();

                ContactEntry googleContact = new ContactEntry();
                ContactSync.UpdateContact(outlookContact, googleContact);
                match = new ContactMatch(outlookContact, googleContact);

                //save contact to google.
                sync.SaveGoogleContact(match);
            }

            sync.Load();
            ContactsMatcher.SyncContacts(sync);

            // all contacts should be synced
            Assert.AreEqual(c, sync.Contacts.Count - excount);

            DateTime end = DateTime.Now;
            TimeSpan time = end - start;
            Console.WriteLine("Synced " + c + " contacts to google in " + time.TotalSeconds + " s ("
                + ((float)time.TotalSeconds / (float)c) + " s per contact)");

            // received: Synced 50 contacts to google in 30.137 s (0.60274 s per contact)
        }

        [Ignore]
        public void TestCreatingGoogeAccountThatFailed1()
        {
            Outlook.ContactItem outlookContact = sync.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"], 
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNull(match.GoogleContact);

            ContactEntry googleContact = new ContactEntry();

            //ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title.Text = outlookContact.FileAs;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.FullName;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            ContactSync.SetAddresses(outlookContact, googleContact);

            ContactSync.SetCompanies(outlookContact, googleContact);

            ContactSync.SetIMs(outlookContact, googleContact);

            googleContact.Content.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            ContactEntry createdEntry = (ContactEntry)sync.GoogleService.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, match.OutlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            match.OutlookContact.Save();
        }

        //[Test]
        [Ignore]
        public void TestCreatingGoogeAccountThatFailed2()
        {
            Outlook.ContactItem outlookContact = sync.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNull(match.GoogleContact);

            ContactEntry googleContact = new ContactEntry();

            //ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title.Text = outlookContact.FileAs;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.FullName;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            //SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(outlookContact, googleContact);

            //SetCompanies(outlookContact, googleContact);

            //SetIMs(outlookContact, googleContact);

            googleContact.Content.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            ContactEntry createdEntry = (ContactEntry)sync.GoogleService.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, match.OutlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            match.OutlookContact.Save();
        }

        //[Test]
        [Ignore]
        public void TestCreatingGoogeAccountThatFailed3()
        {
            Outlook.ContactItem outlookContact = sync.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNull(match.GoogleContact);

            ContactEntry googleContact = new ContactEntry();

            //ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title.Text = outlookContact.FileAs;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.FullName;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.CompanyName;

            //SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(outlookContact, googleContact);

            //SetCompanies(outlookContact, googleContact);

            //SetIMs(outlookContact, googleContact);

            //googleContact.Content.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            ContactEntry createdEntry = (ContactEntry)sync.GoogleService.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, match.OutlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            match.OutlookContact.Save();
        }

        //[Test]
        [Ignore]
        public void TestUpdatingGoogeAccountThatFailed()
        {
            Outlook.ContactItem outlookContact = sync.OutlookContacts.Find(
                string.Format("[FirstName]='{0}' AND [LastName]='{1}'",
                ConfigurationManager.AppSettings["Test.FirstName"],
                ConfigurationManager.AppSettings["Test.LastName"])) as Outlook.ContactItem;

            ContactMatch match = FindMatch(outlookContact);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleContact);

            ContactEntry googleContact = match.GoogleContact;

            ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title.Text = outlookContact.FileAs;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.FullName;

            if (googleContact.Title.Text == null)
                googleContact.Title.Text = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(outlookContact, googleContact);

            //SetCompanies(outlookContact, googleContact);

            //SetIMs(outlookContact, googleContact);

            //googleContact.Content.Content = outlookContact.Body;

            ContactEntry updatedEntry = googleContact.Update() as ContactEntry;

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, match.OutlookContact, updatedEntry);
            match.GoogleContact = updatedEntry;
            match.OutlookContact.Save();
        }

        private void DeleteExistingTestContacts(string name, string email)
        {
            sync.Load();
            ContactsMatcher.SyncContacts(sync);
            ContactMatch match = sync.ContactByProperty(name, email);
            try
            {
                while (match != null)
                {
                    if (match.GoogleContact != null)
                        match.GoogleContact.Delete();
                    if (match.OutlookContact != null)
                        match.OutlookContact.Delete();

                    sync.Load();
                    ContactsMatcher.SyncContacts(sync);
                    match = sync.ContactByProperty(name, email);
                }
            }
            catch { }

            Outlook.ContactItem prevOutlookContact = sync.OutlookContacts.Find("[Email1Address] = '" + email + "'") as Outlook.ContactItem;
            if (prevOutlookContact != null)
                prevOutlookContact.Delete();
        }

        internal ContactMatch FindMatch(Outlook.ContactItem outlookContact)
        {
            foreach (ContactMatch match in sync.Contacts)
            {
                if (match.OutlookContact.EntryID == outlookContact.EntryID)
                    return match;
            }
            return null;
        }
        internal ContactMatch FindMatch(Outlook.ContactItem outlookContact, ContactMatchList matches)
        {
            foreach (ContactMatch match in matches)
            {
                if (match.OutlookContact.EntryID == outlookContact.EntryID)
                    return match;
            }
            return null;
        }
    }
}
