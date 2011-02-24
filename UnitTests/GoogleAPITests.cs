using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using System.Configuration;
using Google.Contacts;

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
                //ContactsService service = new ContactsService("WebGear.GoogleContactsSync");
                //service.setUserCredentials(ConfigurationManager.AppSettings["Gmail.Username"], 
                //    ConfigurationManager.AppSettings["Gmail.Password"]);
                RequestSettings rs = new RequestSettings("GoogleContactSyncMod", ConfigurationManager.AppSettings["Gmail.Username"], ConfigurationManager.AppSettings["Gmail.Password"]);
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
            catch (Exception)
            {

            }
        }
    }
}
