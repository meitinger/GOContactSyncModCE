using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using System.Collections;
using Google.Contacts;

namespace GoContactSyncMod
{
    internal static class ContactPropertiesUtils
    {
        public static string GetOutlookId(Outlook.ContactItem outlookContact)
        {
            return outlookContact.EntryID;
        }
        public static string GetGoogleId(Contact googleContact)
        {
            string id = googleContact.Id.ToString();
            if (id == null)
                throw new Exception();
            return id;
        }

        public static void SetGoogleOutlookContactId(string syncProfile, Contact googleContact, Outlook.ContactItem outlookContact)
        {
            if (outlookContact.EntryID == null)
                throw new Exception("Must save outlook contact before getting id");

            SetGoogleOutlookContactId(syncProfile, googleContact, GetOutlookId(outlookContact));
        }
        public static void SetGoogleOutlookContactId(string syncProfile, Contact googleContact, string outlookContactId)
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
        public static string GetGoogleOutlookContactId(string syncProfile, Contact googleContact)
        {
            // get extended prop
            foreach (Google.GData.Extensions.ExtendedProperty p in googleContact.ExtendedProperties)
            {
                if (p.Name == "gos:oid:" + syncProfile + "")
                    return (string)p.Value;
            }
            return null;
        }
        public static void ResetGoogleOutlookContactId(string syncProfile, Contact googleContact)
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

        /// <summary>
        /// Sets the syncId of the Outlook contact and the last sync date. 
        /// Please assure to always call this function when saving OutlookItem
        /// </summary>
        /// <param name="sync"></param>
        /// <param name="outlookContact"></param>
        /// <param name="googleContact"></param>
        public static void SetOutlookGoogleContactId(Syncronizer sync, Outlook.ContactItem outlookContact, Contact googleContact)
        {
            if (googleContact.ContactEntry.Id.Uri == null)
                throw new NullReferenceException("GoogleContact must have a valid Id");

            //check if outlook contact aready has google id property.
            Outlook.UserProperty prop = outlookContact.UserProperties[sync.OutlookPropertyNameId];
            if (prop == null)
                prop = outlookContact.UserProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, null, null);
            prop.Value = googleContact.ContactEntry.Id.Uri.Content;

            //save last google's updated date as property
            /*prop = outlookContact.UserProperties[OutlookPropertyNameUpdated];
            if (prop == null)
                prop = outlookContact.UserProperties.Add(OutlookPropertyNameUpdated, Outlook.OlUserPropertyType.olDateTime, null, null);
            prop.Value = googleContact.Updated;*/

            //Also set the OutlookLastSync date when setting a match between Outlook and Google to assure the lastSync updated when Outlook contact is saved afterwards
            SetOutlookLastSync(sync, outlookContact);
        }

        public static void SetOutlookLastSync(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            //save sync datetime
            Outlook.UserProperty prop = outlookContact.UserProperties[sync.OutlookPropertyNameSynced];
            if (prop == null)
                prop = outlookContact.UserProperties.Add(sync.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, null, null);
            prop.Value = DateTime.Now;
        }

        public static DateTime? GetOutlookLastSync(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            Outlook.UserProperty prop = outlookContact.UserProperties[sync.OutlookPropertyNameSynced];
            if (prop != null)
                return (DateTime)prop.Value;
            return null;
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

            List<int> indexesToBeRemoved = new List<int>();
            IEnumerator en = outlookContact.UserProperties.GetEnumerator();
            en.Reset();
            int index = 1; // 1 based collection            
            while (en.MoveNext())            
            {
                if (en.Current as Outlook.UserProperty == idProp || en.Current as Outlook.UserProperty == lastSyncProp)
                {
                    indexesToBeRemoved.Add(index);
                    //outlookContact.UserProperties.Remove(index);
                    //Don't return to remove both properties, googleId and lastSynced
                    //return;
                }
                index++;
            }

            for (int i = indexesToBeRemoved.Count-1; i>=0 ; i--)
                outlookContact.UserProperties.Remove(indexesToBeRemoved[i]);
            //throw new Exception("Did not find prop.");
        }
    }
}
