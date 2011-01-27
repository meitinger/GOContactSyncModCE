using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using System.Configuration;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class GoogleAPITests
    {
        [Test]
        public void CreateNewContact()
        {
            try
            {
                ContactsService service = new ContactsService("WebGear.GoogleContactsSync");
                service.setUserCredentials(ConfigurationManager.AppSettings["Gmail.Username"], 
                    ConfigurationManager.AppSettings["Gmail.Password"]);
                
                #region Delete previously created test contact.
                ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
                query.NumberToRetrieve = 500;

                ContactsFeed feed = service.Query(query);

                foreach (ContactEntry entry in feed.Entries)
                {
                    if (entry.PrimaryEmail != null && entry.PrimaryEmail.Address == "johndoe@example.com")
                    {
                        entry.Delete();
                        break;
                    }
                } 
                #endregion
                
                ContactEntry newEntry = new ContactEntry();
                newEntry.Title.Text = "John Doe";
                    
                EMail primaryEmail = new EMail("johndoe@example.com");
                primaryEmail.Primary = true;
                primaryEmail.Rel = ContactsRelationships.IsWork;
                newEntry.Emails.Add(primaryEmail);

                PhoneNumber phoneNumber = new PhoneNumber("555-555-5551");
                phoneNumber.Primary = true;
                phoneNumber.Rel = ContactsRelationships.IsMobile;
                newEntry.Phonenumbers.Add(phoneNumber);

                PostalAddress postalAddress = new PostalAddress();
                postalAddress.Value = "123 somewhere lane";
                postalAddress.Primary = true;
                postalAddress.Rel = ContactsRelationships.IsHome;
                newEntry.PostalAddresses.Add(postalAddress);

                newEntry.Content.Content = "Who is this guy?";

                Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

                ContactEntry createdEntry = (ContactEntry)service.Insert(feedUri, newEntry);

                Assert.IsNotNull(createdEntry.Id.Uri);

                //delete this temp contact
                createdEntry.Delete();
            }
            catch (Exception)
            {

            }
        }
    }
}
