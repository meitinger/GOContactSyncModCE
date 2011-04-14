using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using System.Collections;
using Google.Contacts;
using System.Runtime.InteropServices;

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
            Outlook.UserProperties userProperties = outlookContact.UserProperties;
            try
            {
                Outlook.UserProperty prop = userProperties[sync.OutlookPropertyNameId];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameId, Outlook.OlUserPropertyType.olText, true);
                try
                {
                    prop.Value = googleContact.ContactEntry.Id.Uri.Content;
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }

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
            Outlook.UserProperties userProperties = outlookContact.UserProperties;
            try
            {
                Outlook.UserProperty prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop == null)
                    prop = userProperties.Add(sync.OutlookPropertyNameSynced, Outlook.OlUserPropertyType.olDateTime, true);
                try
                {
                    prop.Value = DateTime.Now;
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static DateTime? GetOutlookLastSync(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            DateTime? result = null;
            Outlook.UserProperties userProperties = outlookContact.UserProperties;
            try
            {
                Outlook.UserProperty prop = userProperties[sync.OutlookPropertyNameSynced];
                if (prop != null)
                {
                    try
                    {
                        result = (DateTime)prop.Value;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(prop);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
            return result;
        }
        public static string GetOutlookGoogleContactId(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            string id = null;
            Outlook.UserProperties userProperties = outlookContact.UserProperties;
            try
            {
                Outlook.UserProperty idProp = userProperties[sync.OutlookPropertyNameId];
                if (idProp != null)
                {
                    try
                    {
                        id = (string)idProp.Value;
                        if (id == null)
                            throw new Exception();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(idProp);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
            return id;
        }
        public static void ResetOutlookGoogleContactId(Syncronizer sync, Outlook.ContactItem outlookContact)
        {
            Outlook.UserProperties userProperties = outlookContact.UserProperties;
            try
            {
                Outlook.UserProperty idProp = userProperties[sync.OutlookPropertyNameId];
                try
                {
                    Outlook.UserProperty lastSyncProp = userProperties[sync.OutlookPropertyNameSynced];
                    try
                    {
                        if (idProp == null && lastSyncProp == null)
                            return;

                        List<int> indexesToBeRemoved = new List<int>();
                        IEnumerator en = userProperties.GetEnumerator();
                        en.Reset();
                        int index = 1; // 1 based collection            
                        while (en.MoveNext())
                        {
                            Outlook.UserProperty userProperty = en.Current as Outlook.UserProperty;
                            if (userProperty == idProp || userProperty == lastSyncProp)
                            {
                                indexesToBeRemoved.Add(index);
                                //outlookContact.UserProperties.Remove(index);
                                //Don't return to remove both properties, googleId and lastSynced
                                //return;
                            }
                            index++;
                            Marshal.ReleaseComObject(userProperty);
                        }

                        for (int i = indexesToBeRemoved.Count - 1; i >= 0; i--)
                            userProperties.Remove(indexesToBeRemoved[i]);
                        //throw new Exception("Did not find prop.");
                    }
                    finally
                    {
                        if (lastSyncProp != null)
                            Marshal.ReleaseComObject(lastSyncProp);
                    }
                }
                finally
                {
                    if (idProp != null)
                        Marshal.ReleaseComObject(idProp);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(userProperties);
            }
        }

        public static string GetOutlookEmailAddress1(Outlook.ContactItem outlookContactItem)
        {
            return GetOutlookEmailAddress(outlookContactItem, outlookContactItem.Email1AddressType, outlookContactItem.Email1EntryID, outlookContactItem.Email1Address);
        }

        public static string GetOutlookEmailAddress2(Outlook.ContactItem outlookContactItem)
        {
            return GetOutlookEmailAddress(outlookContactItem, outlookContactItem.Email2AddressType, outlookContactItem.Email2EntryID, outlookContactItem.Email2Address);
        }

        public static string GetOutlookEmailAddress3(Outlook.ContactItem outlookContactItem)
        {
            return GetOutlookEmailAddress(outlookContactItem, outlookContactItem.Email3AddressType, outlookContactItem.Email3EntryID, outlookContactItem.Email3Address);
        }

        private static string GetOutlookEmailAddress(Outlook.ContactItem outlookContactItem, string emailAddressType, string emailEntryID, string emailAddress)
        {
            switch (emailAddressType)
            {
                case "EX":  // Microsoft Exchange address: "/o=xxxx/ou=xxxx/cn=Recipients/cn=xxxx"
                    Outlook.NameSpace outlookNameSpace = outlookContactItem.Application.GetNamespace("mapi");
                    try
                    {
                        // The emailEntryID is garbage (bug in Outlook 2007 and before?) - so we cannot do GetAddressEntryFromID().
                        // Instead we create a temporary recipient and ask Exchange to resolve it, then get the SMTP address from it.
                        //Outlook.AddressEntry addressEntry = outlookNameSpace.GetAddressEntryFromID(emailEntryID);
                        Outlook.Recipient recipient = outlookNameSpace.CreateRecipient(emailAddress);
                        try
                        {
                            recipient.Resolve();
                            if (recipient.Resolved)
                            {
                                Outlook.AddressEntry addressEntry = recipient.AddressEntry;
                                if (addressEntry != null)
                                {
                                    try
                                    {
                                        if (addressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                                        {
                                            Outlook.ExchangeUser exchangeUser = addressEntry.GetExchangeUser();
                                            if (exchangeUser != null)
                                            {
                                                try
                                                {
                                                    return exchangeUser.PrimarySmtpAddress;
                                                }
                                                finally
                                                {
                                                    Marshal.ReleaseComObject(exchangeUser);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            Logger.Log(string.Format("Unsupported AddressEntryUserType {0} for contact '{1}'.", addressEntry.AddressEntryUserType, outlookContactItem.FileAs), EventType.Debug);
                                        }
                                    }
                                    finally
                                    {
                                        Marshal.ReleaseComObject(addressEntry);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (recipient != null)
                                Marshal.ReleaseComObject(recipient);
                        }
                    }
                    finally
                    {
                        if (outlookNameSpace != null)
                            Marshal.ReleaseComObject(outlookNameSpace);
                    }
                    // Fallback: If Exchange cannot give us the SMTP address, we give up and use the Exchange address format.
                    // TODO: Can we do better?
                    return emailAddress;

                case "SMTP":
                default:
                    return emailAddress;
            }
        }
    }
}
