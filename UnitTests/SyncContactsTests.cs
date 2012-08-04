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
using Google.Contacts;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncContactsTests
    {
        Syncronizer sync;

        static Logger.LogUpdatedHandler _logUpdateHandler = null;

        //Constants for test contact
        const string name = "AN_OUTLOOK_TEST_CONTACT";
        const string email = "email00@outlook.com";
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
            sync.SyncContacts = true;
            sync.SyncNotes = false;
            sync.SyncProfile = syncProfile;
            Syncronizer.SyncContactsFolder = syncContactsFolder;  

            sync.LoginToGoogle(gmailUsername, gmailPassword);
            sync.LoginToOutlook();

        }

        [SetUp]
        public void SetUp()
        {
            // delete previously failed test contacts
            DeleteTestContacts();
            sync.UseFileAs = true;
                      
        }

        private void DeleteTestContacts()
        {
            sync.LoadContacts();

            //foreach (Outlook.ContactItem outlookContact in sync.OutlookContacts)
            //{
            //    if (outlookContact != null && 
            //        ((outlookContact.Email1Address != null && outlookContact.Email1Address == email) ||
            //          outlookContact.FileAs == name))
            //    {
            //            DeleteTestContact(outlookContact);
            //    }
            //}

            Outlook.ContactItem outlookContact = sync.OutlookContacts.Find("[Email1Address] = '" + email + "'") as Outlook.ContactItem;
            if (outlookContact != null)
                DeleteTestContact(outlookContact);

            outlookContact = sync.OutlookContacts.Find("[FileAs] = '" + name + "'") as Outlook.ContactItem;
            if (outlookContact != null)
                DeleteTestContact(outlookContact);

            outlookContact = sync.OutlookContacts.Find("[FileAs] = 'SaveAs'") as Outlook.ContactItem;
            if (outlookContact != null)
                DeleteTestContact(outlookContact);

            foreach (Contact googleContact in sync.GoogleContacts)
            {
                if (googleContact != null &&
                    ((googleContact.PrimaryEmail != null && googleContact.PrimaryEmail.Address == email) ||
                      googleContact.Title == name ||
                      googleContact.Name.FullName == name))
                {
                    DeleteTestContact(googleContact);
                }
            }
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
        public void TestSync_Structured()
        {        

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            
            //outlookContact.HomeAddress = "10 Parades";
            outlookContact.HomeAddressStreet = "Street";
            outlookContact.HomeAddressCity = "City";
            outlookContact.HomeAddressPostalCode = "1234";
            outlookContact.HomeAddressCountry = "Country";
            outlookContact.HomeAddressPostOfficeBox = "PO1";
            outlookContact.HomeAddressState = "State1";
            
            //outlookContact.BusinessAddress = "11 Parades"
            outlookContact.BusinessAddressStreet = "Street2";
            outlookContact.BusinessAddressCity = "City2";
            outlookContact.BusinessAddressPostalCode = "5678";
            outlookContact.BusinessAddressCountry = "Country2";
            outlookContact.BusinessAddressPostOfficeBox = "PO2";
            outlookContact.BusinessAddressState = "State2";

            ///outlookContact.OtherAddress = "12 Parades";
            outlookContact.OtherAddressStreet = "Street3";
            outlookContact.OtherAddressCity = "City3";
            outlookContact.OtherAddressPostalCode = "8012";
            outlookContact.OtherAddressCountry = "Country3";
            outlookContact.OtherAddressPostOfficeBox = "PO3";
            outlookContact.OtherAddressState = "State3";

            #region phones
            //First delete the destination phone numbers
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.HomeTelephoneNumber = "456";
            outlookContact.Home2TelephoneNumber = "4567";
            outlookContact.BusinessTelephoneNumber = "45678";
            outlookContact.Business2TelephoneNumber = "456789";
            outlookContact.MobileTelephoneNumber = "123";
            outlookContact.BusinessFaxNumber = "1234";
            outlookContact.HomeFaxNumber = "12345";
            outlookContact.PagerNumber = "123456";
            //outlookContact.RadioTelephoneNumber = "1234567";
            outlookContact.OtherTelephoneNumber = "12345678";
            outlookContact.CarTelephoneNumber = "123456789";
            outlookContact.AssistantTelephoneNumber = "987";
            #endregion phones

            #region Name
            outlookContact.Title = "Title";
            outlookContact.FirstName = "Firstname";            
            outlookContact.MiddleName = "Middlename";
            outlookContact.LastName = "Lastname";
            outlookContact.Suffix = "Suffix";
            //outlookContact.FullName = name; //The Outlook fullName is automatically set, so don't assign it from Google
            #endregion Name

            outlookContact.Birthday = new DateTime(1999,1,1);

            outlookContact.NickName = "Nickname";
            outlookContact.OfficeLocation = "Location";            
            outlookContact.Initials = "IN";
            outlookContact.Language = "German"; 
            
            //outlookContact.Companies = "Company";
            outlookContact.CompanyName = "CompanyName";
            outlookContact.JobTitle = "Position";
            outlookContact.Department = "Department";

            outlookContact.IMAddress = "IMs";
            outlookContact.Anniversary = new DateTime(2000,1,1);
            outlookContact.Children = "Children";
            outlookContact.Spouse = "Spouse";
            outlookContact.AssistantName = "Assi";
            outlookContact.ManagerName = "Chef";
            outlookContact.WebPage = "http://www.test.de";
            outlookContact.Body = "<sn>Content & other stuff</sn>\r\n<![CDATA[  \r\n...\r\n&stuff in CDATA < >\r\n  ]]>";
            outlookContact.Save();
            
            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;     
            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            Outlook.ContactItem recreatedOutlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);
            Assert.AreEqual(outlookContact.Email1Address, recreatedOutlookContact.Email1Address);
            Assert.AreEqual(outlookContact.Email2Address, recreatedOutlookContact.Email2Address);
            Assert.AreEqual(outlookContact.Email3Address, recreatedOutlookContact.Email3Address);
            Assert.AreEqual(outlookContact.PrimaryTelephoneNumber, recreatedOutlookContact.PrimaryTelephoneNumber);
            Assert.AreEqual(outlookContact.HomeTelephoneNumber, recreatedOutlookContact.HomeTelephoneNumber);
            Assert.AreEqual(outlookContact.Home2TelephoneNumber, recreatedOutlookContact.Home2TelephoneNumber);
            Assert.AreEqual(outlookContact.BusinessTelephoneNumber, recreatedOutlookContact.BusinessTelephoneNumber);
            Assert.AreEqual(outlookContact.Business2TelephoneNumber, recreatedOutlookContact.Business2TelephoneNumber);
            Assert.AreEqual(outlookContact.MobileTelephoneNumber, recreatedOutlookContact.MobileTelephoneNumber);
            Assert.AreEqual(outlookContact.BusinessFaxNumber, recreatedOutlookContact.BusinessFaxNumber);
            Assert.AreEqual(outlookContact.HomeFaxNumber, recreatedOutlookContact.HomeFaxNumber);
            Assert.AreEqual(outlookContact.PagerNumber, recreatedOutlookContact.PagerNumber);
            //Assert.AreEqual(outlookContact.RadioTelephoneNumber, recreatedOutlookContact.RadioTelephoneNumber);
            Assert.AreEqual(outlookContact.OtherTelephoneNumber, recreatedOutlookContact.OtherTelephoneNumber);
            Assert.AreEqual(outlookContact.CarTelephoneNumber, recreatedOutlookContact.CarTelephoneNumber);
            Assert.AreEqual(outlookContact.AssistantTelephoneNumber, recreatedOutlookContact.AssistantTelephoneNumber);

            Assert.AreEqual(outlookContact.HomeAddressStreet, recreatedOutlookContact.HomeAddressStreet);
            Assert.AreEqual(outlookContact.HomeAddressCity, recreatedOutlookContact.HomeAddressCity);
            Assert.AreEqual(outlookContact.HomeAddressCountry, recreatedOutlookContact.HomeAddressCountry);
            Assert.AreEqual(outlookContact.HomeAddressPostalCode, recreatedOutlookContact.HomeAddressPostalCode);
            Assert.AreEqual(outlookContact.HomeAddressPostOfficeBox, recreatedOutlookContact.HomeAddressPostOfficeBox);
            Assert.AreEqual(outlookContact.HomeAddressState, recreatedOutlookContact.HomeAddressState);

            Assert.AreEqual(outlookContact.BusinessAddressStreet, recreatedOutlookContact.BusinessAddressStreet);
            Assert.AreEqual(outlookContact.BusinessAddressCity, recreatedOutlookContact.BusinessAddressCity);
            Assert.AreEqual(outlookContact.BusinessAddressCountry, recreatedOutlookContact.BusinessAddressCountry);
            Assert.AreEqual(outlookContact.BusinessAddressPostalCode, recreatedOutlookContact.BusinessAddressPostalCode);
            Assert.AreEqual(outlookContact.BusinessAddressPostOfficeBox, recreatedOutlookContact.BusinessAddressPostOfficeBox);
            Assert.AreEqual(outlookContact.BusinessAddressState, recreatedOutlookContact.BusinessAddressState);

            Assert.AreEqual(outlookContact.OtherAddressStreet, recreatedOutlookContact.OtherAddressStreet);
            Assert.AreEqual(outlookContact.OtherAddressCity, recreatedOutlookContact.OtherAddressCity);
            Assert.AreEqual(outlookContact.OtherAddressCountry, recreatedOutlookContact.OtherAddressCountry);
            Assert.AreEqual(outlookContact.OtherAddressPostalCode, recreatedOutlookContact.OtherAddressPostalCode);
            Assert.AreEqual(outlookContact.OtherAddressPostOfficeBox, recreatedOutlookContact.OtherAddressPostOfficeBox);
            Assert.AreEqual(outlookContact.OtherAddressState, recreatedOutlookContact.OtherAddressState);

            Assert.AreEqual(outlookContact.FullName, recreatedOutlookContact.FullName);
            Assert.AreEqual(outlookContact.MiddleName, recreatedOutlookContact.MiddleName);
            Assert.AreEqual(outlookContact.LastName, recreatedOutlookContact.LastName);
            Assert.AreEqual(outlookContact.FirstName, recreatedOutlookContact.FirstName);
            Assert.AreEqual(outlookContact.Title, recreatedOutlookContact.Title);
            Assert.AreEqual(outlookContact.Suffix, recreatedOutlookContact.Suffix);

            Assert.AreEqual(outlookContact.Birthday, recreatedOutlookContact.Birthday);

            Assert.AreEqual(outlookContact.NickName, recreatedOutlookContact.NickName);
            Assert.AreEqual(outlookContact.OfficeLocation, recreatedOutlookContact.OfficeLocation);
            Assert.AreEqual(outlookContact.Initials, recreatedOutlookContact.Initials);
            Assert.AreEqual(outlookContact.Language, recreatedOutlookContact.Language);

            Assert.AreEqual(outlookContact.IMAddress, recreatedOutlookContact.IMAddress); 
            Assert.AreEqual(outlookContact.Anniversary, recreatedOutlookContact.Anniversary); 
            Assert.AreEqual(outlookContact.Children, recreatedOutlookContact.Children); 
            Assert.AreEqual(outlookContact.Spouse, recreatedOutlookContact.Spouse);
            Assert.AreEqual(outlookContact.ManagerName, recreatedOutlookContact.ManagerName);
            Assert.AreEqual(outlookContact.AssistantName, recreatedOutlookContact.AssistantName); 

            Assert.AreEqual(outlookContact.WebPage, recreatedOutlookContact.WebPage); 
            Assert.AreEqual(outlookContact.Body, recreatedOutlookContact.Body); 

            //Assert.AreEqual(outlookContact.Companies, recreatedOutlookContact.Companies); 
            Assert.AreEqual(outlookContact.CompanyName, recreatedOutlookContact.CompanyName); 
            Assert.AreEqual(outlookContact.JobTitle, recreatedOutlookContact.JobTitle); 
            Assert.AreEqual(outlookContact.Department, recreatedOutlookContact.Department); 

            DeleteTestContacts(match);
            recreatedOutlookContact.Delete();

            //Also delete the birthday/anniversary entries in Outlook calendar
            Logger.Log("Deleting Outlook calendar TEST entries (birthday, anniversary) ...", EventType.Information);

            
            try
            {   Outlook.NameSpace outlookNamespace = Syncronizer.OutlookApplication.GetNamespace("mapi");
                Outlook.MAPIFolder calendarFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                Outlook.Items outlookCalendarItems = calendarFolder.Items;
                for (int i = outlookCalendarItems.Count ; i > 0; i--)
                {
                    Outlook.AppointmentItem item = outlookCalendarItems[i] as Outlook.AppointmentItem;
                    if (item.Subject.Contains(name))
                    {
                        string subject = item.Subject;
                        item.Delete();
                        Logger.Log("Deleted Outlook calendar TEST entry: " + subject, EventType.Information);
                    }

                }
            }
            catch (System.Exception ex)
            {
                Logger.Log("Could not delete Outlook calender TEST entries: " + ex.Message, EventType.Information);
            }

            
           
        }

        [Test]
        public void TestSync_Unstructured()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FileAs = name;          

            outlookContact.HomeAddress = "10 Parades";

            outlookContact.BusinessAddress = "11 Parades";         

            outlookContact.OtherAddress = "12 Parades";

            outlookContact.IMAddress = "  "; //Test empty IMAddress

            outlookContact.Email2Address = "  "; //Test empty Email Address

            outlookContact.FullName = name;
           
            outlookContact.Save();


            sync.SyncOption = SyncOption.OutlookToGoogleOnly;     

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;   
            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            Outlook.ContactItem recreatedOutlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);
           
            Assert.AreEqual(outlookContact.HomeAddress, recreatedOutlookContact.HomeAddress);
            Assert.AreEqual(outlookContact.BusinessAddress, recreatedOutlookContact.BusinessAddress);
            Assert.AreEqual(outlookContact.OtherAddress, recreatedOutlookContact.OtherAddress);

            Assert.AreEqual(outlookContact.FullName, recreatedOutlookContact.FullName);
            
            DeleteTestContacts(match);
            recreatedOutlookContact.Delete();
        }

        [Test]
        public void TestSync_CompanyOnly()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.CompanyName = name;
           
            outlookContact.BusinessAddress = "11 Parades";           

            outlookContact.Save();

            Assert.IsNull(outlookContact.FullName);
            Assert.IsNull(outlookContact.Email1Address);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            Outlook.ContactItem recreatedOutlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);

            Assert.AreEqual(outlookContact.CompanyName, recreatedOutlookContact.CompanyName);

            Assert.AreEqual(outlookContact.BusinessAddress, recreatedOutlookContact.BusinessAddress);

            Assert.IsNull(recreatedOutlookContact.FullName);
            Assert.IsNull(outlookContact.Email1Address);

            DeleteTestContacts(match);
            recreatedOutlookContact.Delete();
        }

        [Test]
        public void TestSync_EmailOnly()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FileAs = email;
            outlookContact.Email1Address = email;

            outlookContact.Save();

            Assert.IsNull(outlookContact.FullName);
            Assert.IsNull(outlookContact.CompanyName);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(email, email);
            //ContactsMatcher.SyncContact(match, sync);

            Outlook.ContactItem recreatedOutlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);

            Assert.AreEqual(outlookContact.Email1Address, recreatedOutlookContact.Email1Address);

            Assert.IsNull(recreatedOutlookContact.FullName);
            Assert.IsNull(outlookContact.CompanyName);

            DeleteTestContacts(match);
            recreatedOutlookContact.Delete();
        }

        [Test]
        public void TestSync_UseFileAs()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.UseFileAs = true;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = "SaveAs";

            outlookContact.Save();

            Assert.AreNotEqual(outlookContact.FullName, outlookContact.FileAs);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty("SaveAs", string.Empty);
            //ContactsMatcher.SyncContact(match, sync);

            Outlook.ContactItem recreatedOutlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(recreatedOutlookContact.FileAs, match.GoogleContact.Title);
            Assert.AreEqual(recreatedOutlookContact.FileAs, match.GoogleContact.Name.FullName);
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);

            recreatedOutlookContact.FileAs = name;
            Assert.AreNotEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);
            Assert.AreNotEqual(recreatedOutlookContact.FileAs, match.GoogleContact.Title);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);

            DeleteTestContacts(match);
            recreatedOutlookContact.Delete();
        }

        [Test]
        public void TestSync_UseFullName()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.UseFileAs = false;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = "SaveAs";           

            outlookContact.Save();

            Assert.AreNotEqual(outlookContact.FullName, outlookContact.FileAs);

            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            //ContactsMatcher.SyncContact(match, sync);

            Outlook.ContactItem recreatedOutlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);

            // match recreatedOutlookContact with outlookContact
            Assert.AreEqual(match.GoogleContact.Name.FullName, OutlookContactInfo.GetTitleFirstLastAndSuffix(recreatedOutlookContact).Trim().Replace("  ", " "));
            Assert.AreNotEqual(outlookContact.FileAs, match.GoogleContact.Title);
            Assert.AreNotEqual(outlookContact.FileAs, match.GoogleContact.Name.FullName);
            Assert.AreNotEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);

            recreatedOutlookContact.FileAs = "SaveAs";
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);
            ContactSync.UpdateContact(match.GoogleContact, recreatedOutlookContact, sync.UseFileAs);
            Assert.AreEqual(outlookContact.FileAs, recreatedOutlookContact.FileAs);

            DeleteTestContacts(match);
            recreatedOutlookContact.Delete();
        }

        [Test]
        public void TestExtendedProps()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.UseFileAs = true;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            sync.SaveGoogleContact(match);

            Assert.AreEqual(name, match.GoogleContact.Title);

            // read contact from google
            googleContact = null;
            MatchContacts(sync);
            ContactsMatcher.SyncContacts(sync);

            match = sync.ContactByProperty(name, email);
            //foreach (ContactMatch m in sync.Contacts)
            //{
            //    if (m.GoogleContact != null && (m.GoogleContact.Title == name || m.GoogleContact.Name.FullName == name))
            //    {
            //        googleContact = m.GoogleContact;
            //        break;
            //    }
            //}

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleContact);

            // get extended prop
            Assert.AreEqual(ContactPropertiesUtils.GetOutlookId(outlookContact), ContactPropertiesUtils.GetGoogleOutlookContactId(sync.SyncProfile, match.GoogleContact));

            DeleteTestContacts(match);    
        }

        [Test]
        public void TestSyncDeletedOulook()
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contacts
            sync.SaveContact(match);

            // delete outlook contact
            outlookContact.Delete();

            // sync
            MatchContacts(sync);
            ContactsMatcher.SyncContacts(sync);
            // find match
            match = null;
            match = sync.ContactByProperty(name, email);
            //foreach (ContactMatch m in sync.Contacts)
            //{
            //    if (m.GoogleContact.Title == name)
            //    {
            //        match = m;
            //        break;
            //    }
            //}

            Assert.IsNotNull(match);

            // delete
            sync.SaveContact(match);

            // sync
            MatchContacts(sync);
            ContactsMatcher.SyncContacts(sync);

            // check if google contact still exists
            googleContact = null;
            match = sync.ContactByProperty(name, email);
            //foreach (ContactMatch m in sync.Contacts)
            //{
            //    if (m.GoogleContact.Title == name)
            //    {
            //        googleContact = m.GoogleContact;
            //        break;
            //    }
            //}
            Assert.IsNull(match);
        }

        [Test]
        public void TestSyncDeletedGoogle()
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;            
            sync.SyncDelete = true;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contacts
            sync.SaveContact(match);

            // delete google contact
            sync.ContactsRequest.Delete(match.GoogleContact);    

            // sync
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // delete
            sync.SaveContact(match);

            // sync
            MatchContacts(sync);
            ContactsMatcher.SyncContacts(sync);
            match = sync.ContactByProperty(name, email);            

            // check if outlook contact still exists
            Assert.IsNull(match);

            DeleteTestContacts(match);    
        }

        [Test]
        public void TestGooglePhoto()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveGoogleContact(match);
            googleContact = null;

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            Image pic = Utilities.CropImageGoogleFormat(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));
            bool saved = Utilities.SaveGooglePhoto(sync, match.GoogleContact, pic);
            Assert.IsTrue(saved);

            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            Image image = Utilities.GetGooglePhoto(sync, match.GoogleContact);
            Assert.IsNotNull(image);

            DeleteTestContacts(match);      
        }

        [Test]
        public void TestOutlookPhoto()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            bool saved = Utilities.SetOutlookPhoto(outlookContact, AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            Assert.IsTrue(saved);
            
            outlookContact.Save();
        
            Image image = Utilities.GetOutlookPhoto(outlookContact);
            Assert.IsNotNull(image);

            outlookContact.Delete();       
        }

        [Test]
        public void TestSyncPhoto()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            Assert.IsTrue(File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg"));
           
            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            Utilities.SetOutlookPhoto(outlookContact, AppDomain.CurrentDomain.BaseDirectory + "\\SamplePic.jpg");
            outlookContact.Save();

            // outlook contact should now have a photo
            Assert.IsNotNull(Utilities.GetOutlookPhoto(outlookContact));

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have a photo
            Assert.IsNotNull(Utilities.GetGooglePhoto(sync, match.GoogleContact));

            // delete outlook contact
            outlookContact.Delete();
            
            // recreate outlook contact
            outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);

            // outlook contact should now have no photo
            Assert.IsNull(Utilities.GetOutlookPhoto(outlookContact));

            sync.UpdateContact(match.GoogleContact, outlookContact);
            match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), match.GoogleContact);
            //match.OutlookContact.Save();            

            //save contact to outlook
            sync.SaveContact(match);

            // outlook contact should now have a photo
            Assert.IsNotNull(Utilities.GetOutlookPhoto(outlookContact));


            Thread.Sleep(5000);
            DeleteTestContacts();                 
        }

        [Test]
        public void TestSyncGroups()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Categories = groupName;
            outlookContact.Save();

            //Outlook contact should now have a group
            Assert.AreEqual(groupName, outlookContact.Categories);

            //Sync Groups first
            MatchContacts(sync);
            ContactsMatcher.SyncGroups(sync);

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //sync and save contact to google.
            ContactsMatcher.SyncContact(match, sync);
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);            

            // google contact should now have the same group
            System.Collections.ObjectModel.Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            Assert.AreEqual(2, googleGroups.Count);
            Assert.Contains(sync.GetGoogleGroupByName(groupName), googleGroups);
            Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myContactsGroup), googleGroups);

            // delete outlook contact
            outlookContact.Delete();
            outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            sync.UpdateContact(match.GoogleContact, outlookContact);
            match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), match.GoogleContact);
            outlookContact.Save();

            sync.SyncOption = SyncOption.MergeGoogleWins;

            //sync and save contact to outlook
            ContactsMatcher.SyncContact(match, sync);
            sync.SaveContact(match);

            //load the same contact from outlook
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);


            Assert.AreEqual(groupName, outlookContact.Categories);

            DeleteTestContacts(match);    

            // delete test group
            Group group = sync.GetGoogleGroupByName(groupName);
            if (group != null)
                sync.ContactsRequest.Delete(group);
        }

        [Test]
        public void TestSyncDeletedGoogleGroup()
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Categories = groupName;
            outlookContact.Save();

            //Outlook contact should now have a group
            Assert.AreEqual(groupName, outlookContact.Categories);

            //Sync Groups first
            MatchContacts(sync);
            ContactsMatcher.SyncGroups(sync);

            //Create now Google Contact and assing new Group
            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.            
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have the same group
            System.Collections.ObjectModel.Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            Group group = sync.GetGoogleGroupByName(groupName);

            Assert.AreEqual(2, googleGroups.Count);
            Assert.Contains(group, googleGroups);
            Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myContactsGroup), googleGroups);

            // delete group from google
            Utilities.RemoveGoogleGroup(match.GoogleContact, group);

            googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            Assert.AreEqual(1, googleGroups.Count);
            Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myContactsGroup), googleGroups);

            //save contact to google.
            sync.SaveGoogleContact(match.GoogleContact);

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;

            //Sync Groups first
            MatchContacts(sync);
            ContactsMatcher.SyncGroups(sync);

            //sync and save contact to outlook.
            match = sync.ContactByProperty(name, email);
            sync.UpdateContact(match.GoogleContact, outlookContact);
            sync.SaveContact(match);            
            
            // google and outlook should now have no category except for the System Group: My Contacts
            googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            Assert.AreEqual(1, googleGroups.Count);
            Assert.AreEqual(null, outlookContact.Categories);
            Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myContactsGroup), googleGroups);

            DeleteTestContacts(match);       
            
            // delete test group
            if (group != null)
                sync.ContactsRequest.Delete(group);
        }

        [Test]
        public void TestSyncDeletedOutlookGroup()
        {
            //ToDo: Check for eache SyncOption and SyncDelete combination
            sync.SyncOption = SyncOption.MergeOutlookWins;
            sync.SyncDelete = true;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Categories = groupName;
            outlookContact.Save();

            //Outlook contact should now have a group
            Assert.AreEqual(groupName, outlookContact.Categories);

            //Now sync Groups
            MatchContacts(sync);            
            ContactsMatcher.SyncGroups(sync);

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should now have the same group
            System.Collections.ObjectModel.Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            Group group = sync.GetGoogleGroupByName(groupName);
            Assert.AreEqual(2, googleGroups.Count);
            Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myContactsGroup),  googleGroups);
            Assert.Contains(group, googleGroups);

            // delete group from outlook
            Utilities.RemoveOutlookGroup(outlookContact, groupName);            
           
            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            sync.UpdateContact(outlookContact, match.GoogleContact);         

            // google and outlook should now have no category
            googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleContact);
            Assert.AreEqual(null, outlookContact.Categories);
            Assert.AreEqual(1, googleGroups.Count);
            Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myContactsGroup),  googleGroups);

            DeleteTestContacts(match);       

            // delete test group
            if (group != null)
                sync.ContactsRequest.Delete(group);
        }

        [Test]
        public void TestResetMatches()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new contact to sync
            Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            //outlookContact.Categories = groupName; //Group is not relevant here
            outlookContact.Save();

            Contact googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            ContactMatch match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // delete outlook contact
            outlookContact.Delete();
            match.OutlookContact = null;

            //load the same contact from google
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            Assert.IsNull(match.OutlookContact);

            // reset matches
            sync.ResetMatch(match.GoogleContact);
            //Not, because NULL: sync.ResetMatch(match.OutlookContact.GetOriginalItemFromOutlook(sync));
            
            // load same contact match
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // google contact should still be present and OutlookContact should be filled
            Assert.IsNotNull(match.GoogleContact);
            Assert.IsNotNull(match.OutlookContact);

            DeleteTestContacts();    

            // create new contact to sync
            outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
            outlookContact.FullName = name;
            outlookContact.FileAs = name;
            outlookContact.Email1Address = email;
            outlookContact.Email2Address = email.Replace("00", "01");
            outlookContact.Email3Address = email.Replace("00", "02");
            outlookContact.HomeAddress = "10 Parades";
            outlookContact.PrimaryTelephoneNumber = "123";
            outlookContact.Save();

            // same test for delete google contact...
            googleContact = new Contact();
            sync.UpdateContact(outlookContact, googleContact);
            match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

            //save contact to google.
            sync.SaveContact(match);

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // delete google contact           
            sync.ContactsRequest.Delete(match.GoogleContact);   
            match.GoogleContact = null;

            //load the same contact from google.
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            Assert.IsNull(match.GoogleContact);

            // reset matches
            //Not, because null: sync.ResetMatch(match.GoogleContact);
            sync.ResetMatch(match.OutlookContact.GetOriginalItemFromOutlook());

            // load same contact match
            MatchContacts(sync);
            match = sync.ContactByProperty(name, email);
            ContactsMatcher.SyncContact(match, sync);

            // Outlook contact should still be present and GoogleContact should be filled
            Assert.IsNotNull(match.OutlookContact);
            Assert.IsNotNull(match.GoogleContact);

            outlookContact.Delete();   
        }

        private void DeleteTestContacts(ContactMatch match)
        {
            if (match != null)
            {
                DeleteTestContact(match.GoogleContact);
                DeleteTestContact(match.OutlookContact);
            }
        }

        private void DeleteTestContact(Outlook.ContactItem outlookContact)
        {
            if (outlookContact != null)
            {
                try
                {
                    string name = outlookContact.FileAs;
                    outlookContact.Delete();
                    Logger.Log("Deleted Outlook test contact: " + name, EventType.Information);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookContact);
                    outlookContact = null;
                }
                
            }
        }


        private void DeleteTestContact(OutlookContactInfo outlookContact)
        {
            if (outlookContact != null)
                DeleteTestContact(outlookContact.GetOriginalItemFromOutlook());             
        }

        private void DeleteTestContact(Contact googleContact)
        {
            if (googleContact != null && !googleContact.Deleted)
            {
                sync.ContactsRequest.Delete(googleContact);
                Logger.Log("Deleted Google test contact: " + googleContact.Title, EventType.Information);
                Thread.Sleep(2000);
            }
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
        //    Outlook.ContactItem outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);
        //    outlookContact.FullName = name;
        //    outlookContact.FileAs = name;
        //    outlookContact.Email1Address = email;
        //    outlookContact.Email2Address = email.Replace("00", "01");
        //    outlookContact.Email3Address = email.Replace("00", "02");
        //    outlookContact.HomeAddress = "10 Parades";
        //    outlookContact.PrimaryTelephoneNumber = "123";
        //    outlookContact.Save();

        //    Contact googleContact = new Contact();
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
                outlookContact = Syncronizer.CreateOutlookContactItem(Syncronizer.SyncContactsFolder);                
                outlookContact.FullName = names[i];
                outlookContact.FileAs = names[i];
                outlookContact.Email1Address = emails[i];
                outlookContact.Save();

                Contact googleContact = new Contact();
                ContactSync.UpdateContact(outlookContact, googleContact, sync.UseFileAs);
                match = new ContactMatch(new OutlookContactInfo(outlookContact, sync), googleContact);

                //save contact to google.
                sync.SaveGoogleContact(match);
            }

            MatchContacts(sync);
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

            Contact googleContact = new Contact();

            //ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            ContactSync.SetAddresses(outlookContact, googleContact);

            ContactSync.SetCompanies(outlookContact, googleContact);

            ContactSync.SetIMs(outlookContact, googleContact);

            googleContact.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdEntry = ((Contact)sync.ContactsRequest.Insert(feedUri, googleContact));
            
            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, outlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            outlookContact.Save();
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

            Contact googleContact = new Contact();

            //ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            //SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(outlookContact, googleContact);

            //SetCompanies(outlookContact, googleContact);

            //SetIMs(outlookContact, googleContact);

            googleContact.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdEntry = (Contact)sync.ContactsRequest.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, outlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            outlookContact.Save();
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

            Contact googleContact = new Contact();

            //ContactSync.UpdateContact(outlookContact, googleContact);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            //SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(outlookContact, googleContact);

            //SetCompanies(outlookContact, googleContact);

            //SetIMs(outlookContact, googleContact);

            //googleContact.Content.Content = outlookContact.Body;

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdEntry = (Contact)sync.ContactsRequest.Insert(feedUri, googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, outlookContact, createdEntry);
            match.GoogleContact = createdEntry;
            outlookContact.Save();
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

            Contact googleContact = match.GoogleContact;

            ContactSync.UpdateContact(outlookContact, googleContact, sync.UseFileAs);

            googleContact.Title = outlookContact.FileAs;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.FullName;

            if (googleContact.Title == null)
                googleContact.Title = outlookContact.CompanyName;

            ContactSync.SetEmails(outlookContact, googleContact);

            ContactSync.SetPhoneNumbers(outlookContact, googleContact);

            //SetAddresses(outlookContact, googleContact);

            //SetCompanies(outlookContact, googleContact);

            //SetIMs(outlookContact, googleContact);

            //googleContact.Content.Content = outlookContact.Body;

            Contact updatedEntry = sync.ContactsRequest.Update(googleContact);

            ContactPropertiesUtils.SetOutlookGoogleContactId(sync, outlookContact, updatedEntry);
            match.GoogleContact = updatedEntry;
            outlookContact.Save();
        }

        //private void DeleteExistingTestContacts(string name, string email)
        //{
        //    MatchContacts(sync);
        //    ContactsMatcher.SyncGroups(sync);
        //    ContactMatch match = sync.ContactByProperty(name, email);

        //    try
        //    {
        //        while (match != null)
        //        {
        //            ContactsMatcher.SyncContact(match, sync);
        //            DeleteTestContacts(match);    

        //            MatchContacts(sync);
        //            match = sync.ContactByProperty(name, email);
        //        }
        //    }
        //    catch { }

        //    Outlook.ContactItem prevOutlookContact = sync.OutlookContacts.Find("[Email1Address] = '" + email + "'") as Outlook.ContactItem;
        //    if (prevOutlookContact != null)
        //        prevOutlookContact.Delete();
        //}

        internal ContactMatch FindMatch(Outlook.ContactItem outlookContact)
        {
            foreach (ContactMatch match in sync.Contacts)
            {
                if (match.OutlookContact.EntryID == outlookContact.EntryID)
                    return match;
            }
            return null;
        }

        private void MatchContacts(Syncronizer sync)
        {
            Thread.Sleep(5000); //Wait, until Contact is really saved and available to retrieve again
            sync.MatchContacts();
        }
        
    }
}
