using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using System.Collections;

namespace GoContactSyncMod
{
    internal static class ContactPropertiesUtils
    {
        public static string GetOutlookId(Outlook.ContactItem outlookContact)
        {
            return outlookContact.EntryID;
        }
        public static string GetGoogleId(ContactEntry googleContact)
        {
            string id = googleContact.Id.ToString();
            if (id == null)
                throw new Exception();
            return id;
        }

        public static void SetGoogleOutlookContactId(string syncProfile, ContactEntry googleContact, Outlook.ContactItem outlookContact)
        {
            if (outlookContact.EntryID == null)
                throw new Exception("Must save outlook contact before getting id");

            SetGoogleOutlookContactId(syncProfile, googleContact, GetOutlookId(outlookContact));
        }
        public static void SetGoogleOutlookContactId(string syncProfile, ContactEntry googleContact, string outlookContactId)
        {
            // check if exists
            bool found = false;
            foreach (Google.GData.Extensions.ExtendedProperty p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                {
                    p.Value = outlookContactId;
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                Google.GData.Extensions.ExtendedProperty prop = new ExtendedProperty(outlookContactId, "gos:oid:" + syncProfile + "");
                prop.Value = outlookContactId;
                googleContact.ExtendedProperties.Add(prop);
            }
        }
        public static string GetGoogleOutlookContactId(string syncProfile, ContactEntry googleContact)
        {
            // get extended prop
            foreach (Google.GData.Extensions.ExtendedProperty p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                    return (string)p.Value;
            }
            return null;
        }
        public static void ResetGoogleOutlookContactId(string syncProfile, ContactEntry googleContact)
        {
            // get extended prop
            foreach (Google.GData.Extensions.ExtendedProperty p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                {
                    // remove 
                    googleContact.ExtendedProperties.Remove(p);
                    return;
                }
            }
        }

        public static void SetOutlookGoogleContactId(Syncronizer sync, Outlook.ContactItem outlookContact, ContactEntry googleContact)
        {
            if (googleContact.Id.Uri == null)
                throw new NullReferenceException("GoogleContact must have a valid Id");

            //check if outlook contact aready has google id property.
            Outlook.UserProperty prop = outlookContact.UserProperties[sync.OutlookPropertyNameId];
            if (prop == null)
                prop = outlookContact.UserProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, null, null);
            prop.Value = googleContact.Id.Uri.Content;

            //save last google's updated date as property
            /*prop = outlookContact.UserProperties[OutlookPropertyNameUpdated];
            if (prop == null)
                prop = outlookContact.UserProperties.Add(OutlookPropertyNameUpdated, Outlook.OlUserPropertyType.olDateTime, null, null);
            prop.Value = googleContact.Updated;*/

            //save sync datetime
            prop = outlookContact.UserProperties[sync.OutlookPropertyNameSynced];
            if (prop == null)
                prop = outlookContact.UserProperties.Add(sync.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, null, null);
            prop.Value = DateTime.Now;
        }
        public static string GetOutlookGoogleContactId(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            Outlook.UserProperty idProp = outlookContact.UserProperties[sync.OutlookPropertyNameId];
            if (idProp == null)
                return null;
            string id = (string)idProp.Value;
            if (id == null)
                throw new Exception();
            return id;
        }
        public static void ResetOutlookGoogleContactId(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            Outlook.UserProperty idProp = outlookContact.UserProperties[sync.OutlookPropertyNameId];
            Outlook.UserProperty lastSyncProp = outlookContact.UserProperties[sync.OutlookPropertyNameSynced];

            if (idProp == null && lastSyncProp == null)
                return;

            IEnumerator en = outlookContact.UserProperties.GetEnumerator();
            int index = 1; // 1 based collection
            while (en.MoveNext())
            {
                if (en.Current as Outlook.UserProperty == idProp || en.Current as Outlook.UserProperty == lastSyncProp)
                {
                    outlookContact.UserProperties.Remove(index);
                    //Don't return to remove both properties, googleId and lastSynced
                    //return;
                }
                index++;
            }

            //throw new Exception("Did not find prop.");
        }
    }
}
