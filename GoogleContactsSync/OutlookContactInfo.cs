using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    /// <summary>
    /// Holds information about an Outlook contact during processing.
    /// We can not always instantiate an unlimited number of Exchange Outlook objects (policy limitations), 
    /// so instead we copy the info we need for our processing into instances of OutlookContactInfo and only
    /// get the real Outlook.ContactItem objects when needed to communicate with Outlook.
    /// </summary>
    class OutlookContactInfo
    {
        #region Internal classes
        internal class UserPropertiesHolder
        {
            public string GoogleContactId;
            public DateTime? LastSync;
        }
        #endregion

        #region Properties
        public string EntryID { get; set; }
        public string FileAs { get; set; }
        public string FullName { get; set; }
        public string Email1Address { get; set; }
        public string MobileTelephoneNumber { get; set; }
        public string Categories { get; set; }
        public DateTime LastModificationTime { get; set; }
        public UserPropertiesHolder UserProperties { get; set; }
        #endregion

        #region Construction
        private OutlookContactInfo()
        {
            // Not public - we are always constructed from an Outlook.ContactItem (constructor below)
        }

        public OutlookContactInfo(ContactItem item, Syncronizer sync)
        {
            this.UserProperties = new UserPropertiesHolder();
            this.Update(item, sync);
        }
        #endregion

        internal void Update(ContactItem outlookContactItem, Syncronizer sync)
        {
            this.EntryID = outlookContactItem.EntryID;
            this.FileAs = outlookContactItem.FileAs;
            this.FullName = outlookContactItem.FullName;
            this.Email1Address = ContactPropertiesUtils.GetOutlookEmailAddress1(outlookContactItem);
            this.MobileTelephoneNumber = outlookContactItem.MobileTelephoneNumber;
            this.Categories = outlookContactItem.Categories;
            this.LastModificationTime = outlookContactItem.LastModificationTime;

            UserProperties userProperties = outlookContactItem.UserProperties;
            UserProperty prop = userProperties[sync.OutlookPropertyNameId];
            this.UserProperties.GoogleContactId = prop != null ? string.Copy((string)prop.Value) : null;
            if (prop != null)
                Marshal.ReleaseComObject(prop);

            prop = userProperties[sync.OutlookPropertyNameSynced];
            this.UserProperties.LastSync = prop != null ? (DateTime)prop.Value : (DateTime?)null;
            if (prop != null)
                Marshal.ReleaseComObject(prop);

            Marshal.ReleaseComObject(userProperties);
        }

        internal ContactItem GetOriginalItemFromOutlook(Syncronizer sync)
        {
            if (this.EntryID == null)
                throw new ApplicationException("OutlookContactInfo cannot re-create the ContactItem from Outlook because EntryID is null, suggesting that this OutlookContactInfo was not created from an existing Outook contact.");

            ContactItem outlookContactItem = sync.OutlookNameSpace.GetItemFromID(this.EntryID) as ContactItem;
            if (outlookContactItem == null)
                throw new ApplicationException("OutlookContactInfo cannot re-create the ContactItem from Outlook because there is no Outlook entry with this EntryID, suggesting that the existing Outook contact may have been deleted.");

            return outlookContactItem;
        }
    }
}
