using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class OutlookTests
    {
        [Test]
        [Ignore("Not needed anymore since it was used as a fix")]
        public void FixUserProperties()
        {
            Outlook.Application outlookApp;
            Outlook.NameSpace outlookNamespace;
            Outlook.Items outlookContacts;

            outlookApp = new Outlook.Application();

            outlookNamespace = outlookApp.GetNamespace("mapi");

            outlookNamespace.Logon("Outlook", null, true, false);

            try
            {
                Outlook.MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                outlookContacts = contactsFolder.Items;

                string oldPrefixId = string.Format("google/contacts/{0}/id", ConfigurationManager.AppSettings["Gmail.Username"]);
                string oldPrefixUpdated = string.Format("google/contacts/{0}/updated", ConfigurationManager.AppSettings["Gmail.Username"]);

                int hashCode = ConfigurationManager.AppSettings["Gmail.Username"].GetHashCode();

                int maxUserIdLength = 32-("g/con/" + "/up").Length;
                string userId = ConfigurationManager.AppSettings["Gmail.Username"];
                if (userId.Length > maxUserIdLength)
                    userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.

                string newPrefixId = "g/con/"+userId+"/id";
                string newPrefixUpdated = "g/con/"+userId+"/up";

                //max property length: 32

                //foreach (Outlook.ContactItem contact in outlookContacts)
                for (int i = 0; i < outlookContacts.Count; i++)
                {
                    try
                    {
                        if (!(outlookContacts[i] is Outlook.ContactItem))
                            continue;
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    Outlook.ContactItem contact = outlookContacts[i] as Outlook.ContactItem;

                    Outlook.UserProperty updatedProp = contact.UserProperties[newPrefixUpdated];
                    if (updatedProp != null)
                    {
                        string lastUpdatedStr = (string)updatedProp.Value;
                        updatedProp.Delete();

                        updatedProp = contact.UserProperties.Add(newPrefixUpdated, Outlook.OlUserPropertyType.olDateTime, null, null);
                        DateTime lastUpdated = DateTime.Parse(lastUpdatedStr);
                        updatedProp.Value = lastUpdated;

                        if (!contact.Saved)
                            contact.Save();

                        continue;
                    }

                    Outlook.UserProperty prop = contact.UserProperties[newPrefixId];
                    if (prop != null)
                        continue;

                    prop = contact.UserProperties[oldPrefixId];
                    if (prop != null)
                    {
                        string id = (string)prop.Value;
                        prop.Delete();

                        prop = contact.UserProperties.Add(newPrefixId, Outlook.OlUserPropertyType.olText, null, null);
                        prop.Value = id;
                    }

                    prop = contact.UserProperties[oldPrefixUpdated];
                    if (prop != null)
                    {
                        DateTime lastUpdated = (DateTime)prop.Value;
                        prop.Delete();

                        prop = contact.UserProperties.Add(newPrefixUpdated, Outlook.OlUserPropertyType.olDateTime, null, null);
                        prop.Value = lastUpdated;
                    }

                    if (!contact.Saved)
                        contact.Save();
                }
            }
            finally
            {
                outlookNamespace.Logoff();
            }
        }
    }
}
