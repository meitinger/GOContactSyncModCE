using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.IO;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Google.Contacts;

namespace GoContactSyncMod
{
	internal class Syncronizer
	{
		public const int OutlookUserPropertyMaxLength = 32;
		public const string OutlookUserPropertyTemplate = "g/con/{0}/";
        internal const string myContactsGroup = "System Group: My Contacts";
		private static object _syncRoot = new object();

		private int _totalCount;
		public int TotalCount
		{
			get { return _totalCount; }
		}

		private int _syncedCount;
		public int SyncedCount
		{
			get { return _syncedCount; }
		}

		private int _deletedCount;
		public int DeletedCount
		{
			get { return _deletedCount; }
		}

        private int _errorCount;
        public int ErrorCount
        {
            get { return _errorCount; }
        }

        private int _skippedCount;
        public int SkippedCount
        {
            set { _skippedCount = value; }
            get { return _skippedCount; }
        }

        private int _skippedCountNotMatches;
        public int SkippedCountNotMatches
        {
            set { _skippedCountNotMatches = value; }
            get { return _skippedCountNotMatches; }
        }


		public delegate void NotificationHandler(string title, string message, EventType eventType);
		public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
		public event NotificationHandler DuplicatesFound;
		public event ErrorNotificationHandler ErrorEncountered;

		private ContactsRequest _googleService;
		public ContactsRequest GoogleService
		{
			get { return _googleService; }
		}

		private Outlook.NameSpace _outlookNamespace;

		private Outlook.Application _outlookApp;
		public Outlook.Application OutlookApplication
		{
			get { return _outlookApp; }
		}

		private Outlook.Items _outlookContacts;
		public Outlook.Items OutlookContacts
		{
			get { return _outlookContacts; }
		}

        private Collection<ContactMatch> _outlookContactDuplicates;
        public Collection<ContactMatch> OutlookContactDuplicates
        {
            get { return _outlookContactDuplicates; }
            set { _outlookContactDuplicates = value; }
        }

        private Collection<ContactMatch> _googleContactDuplicates;
        public Collection<ContactMatch> GoogleContactDuplicates
        {
            get { return _googleContactDuplicates; }
            set { _googleContactDuplicates = value; }
        }

		private Collection<Contact> _googleContacts;
		public Collection<Contact> GoogleContacts
		{
			get { return _googleContacts; }
		}

        private Collection<Group> _googleGroups;
        public Collection<Group> GoogleGroups
		{
			get { return _googleGroups; }
		}

		private string _propertyPrefix;
		public string OutlookPropertyPrefix
		{
			get { return _propertyPrefix; }
		}

		public string OutlookPropertyNameId
		{
			get { return _propertyPrefix + "id"; }
		}

		/*public string OutlookPropertyNameUpdated
		{
			get { return _propertyPrefix + "up"; }
		}*/

		public string OutlookPropertyNameSynced
		{
			get { return _propertyPrefix + "up"; }
		}

		private SyncOption _syncOption = SyncOption.MergeOutlookWins;
		public SyncOption SyncOption
		{
			get { return _syncOption; }
			set { _syncOption = value; }
		}

		private string _syncProfile = "";
		public string SyncProfile
		{
			get { return _syncProfile; }
			set { _syncProfile = value; }
		}

		//private ConflictResolution? _conflictResolution;
		//public ConflictResolution? CResolution
		//{
		//    get { return _conflictResolution; }
		//    set { _conflictResolution = value; }
		//}

		private List<ContactMatch> _matches;
		public List<ContactMatch> Contacts
		{
			get { return _matches; }
		}

        //private string _authToken;
        //public string AuthToken
        //{
        //    get
        //    {
        //        return _authToken;
        //    }
        //}

		private bool _syncDelete;
		/// <summary>
		/// If true deletes contacts if synced before, but one is missing. Otherwise contacts will bever be automatically deleted
		/// </summary>
		public bool SyncDelete
		{
			get { return _syncDelete; }
			set { _syncDelete = value; }
		}


		public Syncronizer()
		{

		}

		public Syncronizer(SyncOption syncOption)
		{
			_syncOption = syncOption;
		}

		public void LoginToGoogle(string username, string password)
		{
			Logger.Log("Connecting to Google...", EventType.Information);
            if (_googleService == null)
            {
                RequestSettings rs = new RequestSettings("GoogleContactSyncMod", username, password); 
                _googleService = new ContactsRequest(rs);
                //_googleService.setUserCredentials(username, password);
                //_authToken = _googleService.QueryClientLoginToken();
            }

			int maxUserIdLength = Syncronizer.OutlookUserPropertyMaxLength - (Syncronizer.OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
			string userId = username;
			if (userId.Length > maxUserIdLength)
				userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.

			_propertyPrefix = string.Format(Syncronizer.OutlookUserPropertyTemplate, userId);
		}

		public void LoginToOutlook()
		{
			Logger.Log("Connecting to Outlook...", EventType.Information);

			try
			{
                CreateOutlookInstance();
			}
			catch (System.Runtime.InteropServices.COMException)
			{
				try
				{
					// If outlook was closed/terminated inbetween, we will receive an Exception
					// System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
					// so recreate outlook instance
					Logger.Log("Cannot connect to Outlook, creating new instance....", EventType.Information);
					/*_outlookApp = new Outlook.Application();
					_outlookNamespace = _outlookApp.GetNamespace("mapi");
					_outlookNamespace.Logon();*/
                    _outlookApp = null;
                    _outlookNamespace = null;
                    CreateOutlookInstance();
                    
				}
				catch (Exception ex)
				{
					string message = String.Format("Cannot connect to Outlook: {0}. \nPlease restart GO Contact Sync Mod and try again. If error persists, please inform developers on SourceForge.", ex.Message);
					// Error again? We need full stacktrace, display it!
					ErrorHandler.Handle(new ApplicationException(message, ex));
				}
			}
		}

        private void CreateOutlookInstance()
        {
            if (_outlookApp == null || _outlookNamespace == null)
            {
                try
                {
                    //Try to create new Outlook application 3 times, because mostly it fails the first time, if not yet running
                    for (int i = 0; i < 3; i++)
                    {
                        try
                        {
                            _outlookApp = new Outlook.Application();
                            break;  //Exit the for loop, if creating outllok application was successful
                        }
                        catch (COMException ex)
                        {
                            if (i == 2)
                                throw ex;
                            else //wait ten seconds and try again
                                System.Threading.Thread.Sleep(1000 * 10);
                        }
                    }

                }
                catch (COMException ex)
                {
                    throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.", ex);
                }

                if (_outlookApp == null)
                    throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");

                try
                {
                    //Try to create new Outlook namespace 3 times, because mostly it fails the first time, if not yet running
                    for (int i = 0; i < 3; i++)
                    {
                        try
                        {
                            _outlookNamespace = _outlookApp.GetNamespace("mapi");
                            break;  //Exit the for loop, if creating outllok application was successful
                        }
                        catch (COMException e)
                        {
                            if (i == 2)
                                throw e;
                            else //wait ten seconds and try again
                                System.Threading.Thread.Sleep(1000 * 10);
                        }
                    }
                   
                }
                catch (COMException comEx)
                {
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", comEx);
                }

                if (_outlookNamespace == null)
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");
            }

            /*
            // Get default profile name from registry, as this is not always "Outlook" and would popup a dialog to choose profile
            // no matter if default profile is set or not. So try to read the default profile, fallback is still "Outlook"
            string profileName = "Outlook";
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Outlook\SocialConnector", false))
            {
                if (k != null)
                    profileName = k.GetValue("PrimaryOscProfile", "Outlook").ToString();
            }
            _outlookNamespace.Logon(profileName, null, true, false);*/

            //Just try to access the outlookNamespace to check, if it is still accessible            
           _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);                                

        }

		public void LogoffOutlook()
		{
            try
            {
                Logger.Log("Disconnecting from Outlook...", EventType.Information);
                if (_outlookNamespace != null)
                {
                    _outlookNamespace.Logoff();
                }
            }
            catch (Exception)
            {
                // if outlook was closed inbetween, we get an System.InvalidCastException or similar exception, that indicates that outlook cannot be acced anymore
                // so as outlook is closed anyways, we just ignore the exception here
            }
            finally
            {
                _outlookNamespace = null;
                _outlookApp = null;
            }
		}

        public void LogoffGoogle()
        {            
            _googleService = null;            
        }

		public void LoadOutlookContacts()
		{
			Logger.Log("Loading Outlook contacts...", EventType.Information);
			Outlook.MAPIFolder contactsFolder = _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            _outlookContacts = contactsFolder.Items;
		}
        ///// <summary>
        ///// Moves duplicates from _outlookContacts to _outlookContactDuplicates
        ///// </summary>
        //private void FilterOutlookContactDuplicates()
        //{
        //    _outlookContactDuplicates = new Collection<Outlook.ContactItem>();
            
        //    if (_outlookContacts.Count < 2)
        //        return;

        //    Outlook.ContactItem main, other;
        //    bool found = true;
        //    int index = 0;

        //    while (found)
        //    {
        //        found = false;

        //        for (int i = index; i <= _outlookContacts.Count - 1; i++)
        //        {
        //            main = _outlookContacts[i] as Outlook.ContactItem;

        //            // only look forward
        //            for (int j = i + 1; j <= _outlookContacts.Count; j++)
        //            {
        //                other = _outlookContacts[j] as Outlook.ContactItem;

        //                if (other.FileAs == main.FileAs &&
        //                    other.Email1Address == main.Email1Address)
        //                {
        //                    _outlookContactDuplicates.Add(other);
        //                    _outlookContacts.Remove(j);
        //                    found = true;
        //                    index = i;
        //                    break;
        //                }
        //            }
        //            if (found)
        //                break;
        //        }
        //    }
        //}

		public void LoadGoogleContacts()
		{
            string message = "Error Loading Google Contacts. Cannot connect to Google: {0}. \r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
			try
			{

				Logger.Log("Loading Google Contacts...", EventType.Information);
                
                _googleContacts = new Collection<Contact>();

				ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
				query.NumberToRetrieve = 256;
				query.StartIndex = 0;

                //Only load Google Contacts in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts")
                Group group = GetGoogleGroupByName(myContactsGroup);
                if (group != null)
                    query.Group = group.Id;

				//query.ShowDeleted = false;
				//query.OrderBy = "lastmodified";
                				
                Feed<Contact> feed=_googleService.Get<Contact>(query);

                while (feed != null)
                {
                    foreach (Contact a in feed.Entries)
                    {
                        _googleContacts.Add(a);
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = _googleService.Get<Contact>(feed, FeedRequestType.Next);
                    
                }                              
	
			}
			catch (System.Net.WebException ex)
			{                               
				//Logger.Log(message, EventType.Error);
                throw new NotSupportedException(string.Format(message, ex.Message), ex);
			}
            catch (System.NullReferenceException ex)
            {                
                //Logger.Log(message, EventType.Error);
                throw new NotSupportedException(string.Format(message, ex.Message), ex);
            }
		}
		public void LoadGoogleGroups()
		{
            string message = "Error Loading Google Groups. Cannot connect to Google: {0}. \r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
            try
            {
                Logger.Log("Loading Google Groups...", EventType.Information);
                GroupsQuery query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;
                //query.ShowDeleted = false;

                _googleGroups = new Collection<Group>();

                Feed<Group> feed = _googleService.Get<Group>(query);               

                while (feed != null)
                {
                    foreach (Group a in feed.Entries)
                    {
                        _googleGroups.Add(a);
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = _googleService.Get<Group>(feed, FeedRequestType.Next);

                }

                ////Only for debugging or reset purpose: Delete all Gougle Groups:
                //for (int i = _googleGroups.Count; i > 0;i-- )
                //    _googleService.Delete(_googleGroups[i-1]);
            }            
			catch (System.Net.WebException ex)
			{                               				
				//Logger.Log(message, EventType.Error);
                throw new NotSupportedException(string.Format(message, ex.Message), ex);
			}
            catch (System.NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new NotSupportedException(string.Format(message, ex.Message), ex);
            }

		}



        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        public void Load()
		{
			LoadOutlookContacts();
            LoadGoogleGroups();
			LoadGoogleContacts();

			DuplicateDataException duplicateDataException;
			_matches = ContactsMatcher.MatchContacts(this, out duplicateDataException);
			if (duplicateDataException != null)
			{
				Logger.Log(duplicateDataException.Message, EventType.Error);
				if (DuplicatesFound != null)
					DuplicatesFound("Google duplicates found", duplicateDataException.Message, EventType.Error);
			}
		}

		public void Sync()
		{
			lock (_syncRoot)
			{
                try
                {
                    if (_syncProfile.Length == 0)
                    {
                        Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                        return;
                    }

                    _syncedCount = 0;
                    _deletedCount = 0;
                    _errorCount = 0;
                    _skippedCount = 0;
                    _skippedCountNotMatches = 0;

                    Load();

#if debug
                    this.DebugContacts();
#endif

                    if (_matches == null)
                        return;

                    _totalCount = _matches.Count + _skippedCountNotMatches;

                    //Remove Google duplicates from matches to be synced
                    if (_googleContactDuplicates != null)
                    {
                        for (int i = _googleContactDuplicates.Count - 1; i >= 0; i--)
                        {
                            ContactMatch match = _googleContactDuplicates[i];
                            if (_matches.Contains(match))
                            {
                                _skippedCount++;
                                _matches.Remove(match);
                            }
                        }
                    }

                    //Remove Outlook duplicates from matches to be synced
                    if (_outlookContactDuplicates != null)
                    {
                        for (int i = _outlookContactDuplicates.Count - 1; i >= 0; i--)
                        {
                            ContactMatch match = _outlookContactDuplicates[i];
                            if (_matches.Contains(match))
                            {
                                _skippedCount++;
                                _matches.Remove(match);
                            }
                        }
                    }

                    ////Remove remaining google contacts not in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts"
                    //Group syncGroup = GetGoogleGroupByName(myContactsGroup);
                    //if (syncGroup != null)
                    //{
                    //    for (int i = _googleContacts.Count -1 ;i >=0; i--)
                    //    {
                    //        Contact googleContact = _googleContacts[i];
                    //        Collection<Group> googleContactGroups = Utilities.GetGoogleGroups(this, googleContact);

                    //        if (!googleContactGroups.Contains(syncGroup))
                    //            _googleContacts.Remove(googleContact);

                    //    }
                    //}                                    


                    Logger.Log("Syncing groups...", EventType.Information);
                    ContactsMatcher.SyncGroups(this);

                    Logger.Log("Syncing contacts...", EventType.Information);
                    ContactsMatcher.SyncContacts(this);

                    SaveContacts(_matches);
                }
                finally
                {              
                    _googleContacts = null;
                    _outlookContacts = null;
                    _outlookContactDuplicates = null;
                    _googleContactDuplicates = null;
                    _googleGroups = null;
                    _matches = null;                    
                }
			}
		}

		public void SaveContacts(List<ContactMatch> contacts)
		{
			foreach (ContactMatch match in contacts)
			{
				try
				{
					SaveContact(match);
				}
				catch (Exception ex)
				{
					if (ErrorEncountered != null)
					{
                        _errorCount++;
                        _syncedCount--;
                        string message = String.Format("Failed to synchronize contact: {0}. \nPlease check the contact, if any Email already exists on Google contacts side or if there is too much or invalid data in the notes field. \nIf the problem persists, please try recreating the contact or report the error on SourceForge.", match.OutlookContact.FullNameAndCompany);
						Exception newEx = new Exception(message, ex);
						ErrorEncountered("Error", newEx, EventType.Error);
					}
					else
						throw;
				}
			}
		}
		public void SaveContact(ContactMatch match)
		{
			if (match.GoogleContact != null && match.OutlookContact != null)
			{
				//bool googleChanged, outlookChanged;
				//SaveContactGroups(match, out googleChanged, out outlookChanged);
                if (match.GoogleContact.ContactEntry.Dirty || match.GoogleContact.ContactEntry.IsDirty() || !match.OutlookContact.Saved)
                    _syncedCount++;

                if (match.GoogleContact.ContactEntry.Dirty || match.GoogleContact.ContactEntry.IsDirty())// || googleChanged)
				{
					//google contact was modified. save.
					SaveGoogleContact(match);
					Logger.Log("Updated Google contact from Outlook: \"" + match.GoogleContact.Title + "\".", EventType.Information);
				}

				if (!match.OutlookContact.Saved)// || outlookChanged)
				{
                    //outlook contact was modified. save.
                    SaveOutlookContact(match);
					Logger.Log("Updated Outlook contact from Google: \"" + match.OutlookContact.FileAs + "\".", EventType.Information);
				}                

				// save photos
				//SaveContactPhotos(match);
			}
			else if (match.GoogleContact == null && match.OutlookContact != null)
			{
				if (ContactPropertiesUtils.GetOutlookGoogleContactId(this, match.OutlookContact) != null)
				{
                    string name = match.OutlookContact.FileAs;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!_syncDelete)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer google contact was deleted, delete outlook contact
                        match.OutlookContact.Delete();
                        _deletedCount++;
                        Logger.Log("Deleted Outlook contact: \"" + name + "\".", EventType.Information);
                    }
				}
			}
			else if (match.GoogleContact != null && match.OutlookContact == null)
			{
				if (ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, match.GoogleContact) != null)
				{
                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because of SyncOption " + _syncOption + ":" + match.GoogleContact.Title + ".", EventType.Information);
                    }
                    else if (!_syncDelete)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because SyncDeletion is switched off :" + match.GoogleContact.Title + ".", EventType.Information);
                    }
                    else
                    {
                        // peer outlook contact was deleted, delete google contact
                        _googleService.Delete(match.GoogleContact);
                        _deletedCount++;
                        Logger.Log("Deleted Google contact: \"" + match.GoogleContact.Title + "\".", EventType.Information);
                    }
				}
			}
			else
			{
				//TODO: ignore for now: throw new ArgumentNullException("To save contacts both ContactMatch peers must be present.");
				Logger.Log("Both Google and Outlook contact: \"" + match.GoogleContact.Title + "\" have been changed! Not implemented yet.", EventType.Warning);
			}
		}

        private void SaveOutlookContact(ContactMatch match)
        {
            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
            match.OutlookContact.Save();
            ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

            Contact updatedEntry = SaveGoogleContact(match.GoogleContact);
            //try
            //{
            //    updatedEntry = _googleService.Update(match.GoogleContact);
            //}
            //catch (GDataRequestException tmpEx)
            //{
            //    // check if it's the known HTCData problem, or if there is any invalid XML element or any unescaped XML sequence
            //    //if (tmpEx.ResponseString.Contains("HTCData") || tmpEx.ResponseString.Contains("&#39") || match.GoogleContact.Content.Contains("<"))
            //    //{
            //    //    bool wasDirty = match.GoogleContact.ContactEntry.Dirty;
            //    //    // XML escape the content
            //    //    match.GoogleContact.Content = EscapeXml(match.GoogleContact.Content);
            //    //    // set dirty to back, cause we don't want the changed content go back to Google without reason
            //    //    match.GoogleContact.ContactEntry.Content.Dirty = wasDirty;
            //    //    updatedEntry = _googleService.Update(match.GoogleContact);
                    
            //    //}
            //    //else 
            //    if (!String.IsNullOrEmpty(tmpEx.ResponseString))
            //        throw new ApplicationException(tmpEx.ResponseString, tmpEx);
            //    else
            //        throw;
            //}            
            match.GoogleContact = updatedEntry;

            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
            match.OutlookContact.Save();
            SaveOutlookPhoto(match);
        }
		private string EscapeXml(string xml)
		{
			string encodedXml = System.Security.SecurityElement.Escape(xml);
			return encodedXml;
		}
		public void SaveGoogleContact(ContactMatch match)
		{
            ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);
            match.GoogleContact = SaveGoogleContact(match.GoogleContact);
            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
            match.OutlookContact.Save();

            //Now save the Photo
            SaveGooglePhoto(match);

		}

		private string GetContactXml(Contact contact)
		{
			MemoryStream ms = new MemoryStream();
			contact.ContactEntry.SaveToXml(ms);
			StreamReader sr = new StreamReader(ms);
			ms.Seek(0, SeekOrigin.Begin);
			string xml = sr.ReadToEnd();
			return xml;
		}

        /// <summary>
        /// Only save the google contact without photo update
        /// </summary>
        /// <param name="googleContact"></param>
		internal Contact SaveGoogleContact(Contact googleContact)
		{
			//check if this contact was not yet inserted on google.
			if (googleContact.ContactEntry.Id.Uri == null)
			{
				//insert contact.
				Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

				try
				{
					Contact createdEntry = _googleService.Insert(feedUri, googleContact);
                    return createdEntry;
				}
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetContactXml(googleContact);
                    string newEx = String.Format("Error saving NEW Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
			}
			else
			{
				try
				{
					//contact already present in google. just update
					
                    // User can create an empty label custom field on the web, but when I retrieve, and update, it throws this:
                    // Data Request Error Response: [Line 12, Column 44, element gContact:userDefinedField] Missing attribute: &#39;key&#39;
                    // Even though I didn't touch it.  So, I will search for empty keys, and give them a simple name.  Better than deleting...
                    int fieldCount = 0;
                    foreach (UserDefinedField userDefinedField in googleContact.ContactEntry.UserDefinedFields)
                    {
                        fieldCount++;
                        if (String.IsNullOrEmpty(userDefinedField.Key))
                        {
                            userDefinedField.Key = "UserField" + fieldCount.ToString();
                        }
                    }

                    //TODO: this will fail if original contact had an empty name or rpimary email address.
                    Contact updated = _googleService.Update(googleContact);
                    return updated;
				}
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetContactXml(googleContact);
                    string newEx = String.Format("Error saving EXISTING Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
			}
		}         

        //public void SaveContactPhotos(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (!hasGooglePhoto && !hasOutlookPhoto)
        //        return;
        //    else if (hasGooglePhoto && _syncOption != SyncOption.OutlookToGoogleOnly)
        //    {
        //        // add google photo to outlook
        //        Image googlePhoto = Utilities.GetGooglePhoto(this, match.GoogleContact);
        //        Utilities.SetOutlookPhoto(match.OutlookContact, googlePhoto);
        //        match.OutlookContact.Save();

        //        googlePhoto.Dispose();
        //    }
        //    else if (hasOutlookPhoto && _syncOption != SyncOption.GoogleToOutlookOnly)
        //    {
        //        // add outlook photo to google
        //        Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
        //        if (outlookPhoto != null)
        //        {
        //            outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
        //            bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
        //            if (!saved)
        //                throw new Exception("Could not save");

        //            outlookPhoto.Dispose();
        //        }
        //    }
        //    else
        //    {
        //        // TODO: if both contacts have photos and one is updated, the
        //        // other will not be updated.
        //    }

        //    //Utilities.DeleteTempPhoto();
        //}

        public void SaveGooglePhoto(ContactMatch match)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

            if (hasOutlookPhoto)
            {
                // add outlook photo to google
                Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);

                if (outlookPhoto != null)
                {
                    using (MemoryStream stream = new MemoryStream(Utilities.BitmapToBytes(new Bitmap(outlookPhoto))))
                    {
                        // Save image to stream.
                        //outlookPhoto.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);

                        //Don'T crop, because maybe someone wants to keep his photo like it is on Outlook
                        //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
                        _googleService.SetPhoto(match.GoogleContact, stream);

                        //Just save the Outlook Contact to have the same lastUpdate date as Google
                        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
                        match.OutlookContact.Save();
                        outlookPhoto.Dispose();
                        
                    }
                }
            }
            else if (hasGooglePhoto)
            {
                //Delete Photo on Google side, if no Outlook photo exists
                _googleService.Delete(match.GoogleContact.PhotoUri, match.GoogleContact.PhotoEtag);
            }

            Utilities.DeleteTempPhoto();
        }

        //public void SaveOutlookPhoto(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (hasGooglePhoto)
        //    {
        //        Image image = new Image(match.GoogleContact.PhotoUri);
        //        Utilities.SetOutlookPhoto(match.OutlookContact, image);
        //        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //        match.OutlookContact.Save();

        //        //googlePhoto.Dispose();
        //    }
        //    else if (hasOutlookPhoto)
        //    {
        //        match.OutlookContact.RemovePicture();
        //        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //        match.OutlookContact.Save();
        //    }
        //}

        //public void SaveGooglePhoto(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (hasOutlookPhoto)
        //    {
        //        // add outlook photo to google
        //        Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
        //        if (outlookPhoto != null)
        //        {
        //            //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
        //            bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
        //            if (!saved)
        //                throw new Exception("Could not save");

        //            //Just save the Outlook Contact to have the same lastUpdate date as Google
        //            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //            match.OutlookContact.Save();
        //            outlookPhoto.Dispose();
        //        }
        //    }
        //    else if (hasGooglePhoto)
        //    {
        //        //ToDo: Delete Photo on Google side, if no Outlook photo exists
        //        //match.GoogleContact.PhotoUri = null;
        //    }

        //    //Utilities.DeleteTempPhoto();
        //}

        public void SaveOutlookPhoto(ContactMatch match)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

            if (hasGooglePhoto)
            {
                // add google photo to outlook
                //ToDo: add google photo to outlook with new Google API
                //Stream stream = _googleService.GetPhoto(match.GoogleContact);
                Image googlePhoto = Utilities.GetGooglePhoto(this, match.GoogleContact);
                Utilities.SetOutlookPhoto(match.OutlookContact, googlePhoto);
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
                match.OutlookContact.Save();

                googlePhoto.Dispose();
            }
            else if (hasOutlookPhoto)
            {
                match.OutlookContact.RemovePicture();
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
                match.OutlookContact.Save();
            }
        }

	
		public Group SaveGoogleGroup(Group group)
		{
			//check if this group was not yet inserted on google.
			if (group.GroupEntry.Id.Uri == null)
			{
				//insert group.
				Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));

				try
				{
					return _googleService.Insert(feedUri, group);
				}
				catch
				{
					//TODO: save google group xml for diagnistics
					throw;
				}
			}
			else
			{
				try
				{
					//group already present in google. just update
					return _googleService.Update(group);
				}
				catch
				{
					//TODO: save google group xml for diagnistics
					throw;
				}
			}
		}

        /// <summary>
        /// Updates Google contact from Outlook (including groups/categories)
        /// </summary>
        public void UpdateContact(Outlook.ContactItem master, Contact slave)
        {
            ContactSync.UpdateContact(master, slave);
            OverwriteContactGroups(master, slave);
        }

        /// <summary>
        /// Updates Outlook contact from Google (including groups/categories)
        /// </summary>
        public void UpdateContact(Contact master, Outlook.ContactItem slave)
        {
            ContactSync.UpdateContact(master, slave);
            OverwriteContactGroups(master, slave);
        }

		/// <summary>
		/// Updates Google contact's groups from Outlook contact
		/// </summary>
		private void OverwriteContactGroups(Outlook.ContactItem master, Contact slave)
		{
			Collection<Group> currentGroups = Utilities.GetGoogleGroups(this, slave);

			// get outlook categories
			string[] cats = Utilities.GetOutlookGroups(master);

			// remove obsolete groups
			Collection<Group> remove = new Collection<Group>();
			bool found;
			foreach (Group group in currentGroups)
			{
				found = false;
				foreach (string cat in cats)
				{
					if (group.Title == cat)
					{
						found = true;
						break;
					}
				}
				if (!found)
					remove.Add(group);
			}
			while (remove.Count != 0)
			{
				Utilities.RemoveGoogleGroup(slave, remove[0]);
				remove.RemoveAt(0);
			}

			// add new groups
			Group g;
			foreach (string cat in cats)
			{
				if (!Utilities.ContainsGroup(this, slave, cat))
				{
					// add group to contact
					g = GetGoogleGroupByName(cat);
					if (g == null)
					{
						throw new Exception(string.Format("Google Groups were supposed to be created prior to saving", cat));
					}
					Utilities.AddGoogleGroup(slave, g);
				}
			}

            //add system Group My Contacts            
            if (!Utilities.ContainsGroup(this, slave, myContactsGroup))
            {
                // add group to contact
                g = GetGoogleGroupByName(myContactsGroup);
                if (g == null)
                {
                    throw new Exception(string.Format("Google System Group: My Contacts doesn't exist", myContactsGroup));
                }
                Utilities.AddGoogleGroup(slave, g);
            }
		}

		/// <summary>
		/// Updates Outlook contact's categories (groups) from Google groups
		/// </summary>
		private void OverwriteContactGroups(Contact master, Outlook.ContactItem slave)
		{
			Collection<Group> newGroups = Utilities.GetGoogleGroups(this, master);

			List<string> newCats = new List<string>(newGroups.Count);
			foreach (Group group in newGroups)
            {   //Only add groups that are no SystemGroup (e.g. "System Group: Meine Kontakte") automatically tracked by Google
                if (group.Title != null && !group.Title.Equals(myContactsGroup))
				    newCats.Add(group.Title);
			}

			slave.Categories = string.Join(", ", newCats.ToArray());
		}

		/// <summary>
		/// Resets associantions of Outlook contacts with Google contacts via user props
		/// and resets associantions of Google contacts with Outlook contacts via extended properties.
		/// </summary>
		public void ResetMatches()
		{
			Debug.Assert(Contacts != null, "Contacts object is null - this should not happen. Please inform Developers.");

            if (_syncProfile.Length == 0)
            {
                Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                return;
            }

            Logger.Log("Resetting matches...", EventType.Information);

			lock (_syncRoot)
			{
				foreach (ContactMatch match in Contacts)
				{
                    try
                    {
                        ResetMatch(match);
                    }
                    catch
                    {
                        Logger.Log("The match of contact " + ((match.OutlookContact!=null)?match.OutlookContact.FileAs:match.GoogleContact.Title) + " couldn't be reset", EventType.Warning);
                    }
				}

				Debug.Assert(Contacts != null, "Contacts object is null after reset - this should not happen. Please inform Developers.");
				Contacts.Clear();
			}
		}

        /// <summary>
        /// Reset the match link between Google and Outlook contact
        /// </summary>
        /// <param name="match"></param>
		public void ResetMatch(ContactMatch match)
		{           
			if (match == null)
				throw new ArgumentNullException("match", "Given ContactMatch is null");
            

            if (match.GoogleContact != null)
            {
                ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, match.GoogleContact);
                SaveGoogleContact(match.GoogleContact);
			}
            
            if (match.OutlookContact != null)
            {
                ContactPropertiesUtils.ResetOutlookGoogleContactId(this, match.OutlookContact);
				match.OutlookContact.Save();
              
                //Reset also Google duplicatesC
                foreach (Contact duplicate in match.AllGoogleContactMatches)
                {
                    if (duplicate != match.GoogleContact)
                    { //To save time, only if not match.GoogleContact, because this was already reset above
                        ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, duplicate);
                        SaveGoogleContact(duplicate);
                    }
                }
			}

            
		}

		public ContactMatch ContactByProperty(string name, string email)
		{
			foreach (ContactMatch m in Contacts)
			{
				if (m.GoogleContact != null &&
					((m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email) ||
					m.GoogleContact.Title == name))
				{
					return m;
				}
				else if (m.OutlookContact != null && (
					(m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email) ||
					m.OutlookContact.FileAs == name))
				{
					return m;
				}
			}
			return null;
		}
		public ContactMatch ContactEmail(string email)
		{
			foreach (ContactMatch m in Contacts)
			{
				if (m.GoogleContact != null &&
					(m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email))
				{
					return m;
				}
				else if (m.OutlookContact != null && (
					m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email))
				{
					return m;
				}
			}
			return null;
		}

		/// <summary>
		/// Used to find duplicates.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public Collection<Outlook.ContactItem> OutlookContactByProperty(string name, string value)
		{
			Collection<Outlook.ContactItem> col = new Collection<Outlook.ContactItem>();
            //foreach (Outlook.ContactItem outlookContact in OutlookContacts)
            //{
            //    if (outlookContact != null && (
            //        (outlookContact.Email1Address != null && outlookContact.Email1Address == email) ||
            //        outlookContact.FileAs == name))
            //    {
            //        col.Add(outlookContact);
            //    }
            //}
            Outlook.ContactItem item = null;
            try
            {
                item = OutlookContacts.Find("["+name+"] = \"" + value + "\"") as Outlook.ContactItem;
                if (item != null)
                {
                    col.Add(item);
                    do
                    {
                        item = OutlookContacts.FindNext() as Outlook.ContactItem;
                        if (item != null)
                            col.Add(item);
                    } while (item != null);
                }
            }
            catch (Exception)
			{
				//TODO: should not get here.
			}

			return col;
		}
		/// <summary>
		/// Used to find duplicates.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="email"></param>
		/// <returns></returns>
		public Collection<Outlook.ContactItem> OutlookContactByEmail(string email)
		{
			Collection<Outlook.ContactItem> col = new Collection<Outlook.ContactItem>();
			Outlook.ContactItem item = null;
			try
			{
				item = OutlookContacts.Find("[Email1Address] = \"" + email + "\"") as Outlook.ContactItem;
				if (item != null)
				{
					col.Add(item);
					do
					{
						item = OutlookContacts.FindNext() as Outlook.ContactItem;
						if (item != null)
							col.Add(item);
					} while (item != null);
				}

                item = OutlookContacts.Find("[Email2Address] = \"" + email + "\"") as Outlook.ContactItem;
                if (item != null)
                {
                    col.Add(item);
                    do
                    {
                        item = OutlookContacts.FindNext() as Outlook.ContactItem;
                        if (item != null)
                            col.Add(item);
                    } while (item != null);
                }

                item = OutlookContacts.Find("[Email3Address] = \"" + email + "\"") as Outlook.ContactItem;
                if (item != null)
                {
                    col.Add(item);
                    do
                    {
                        item = OutlookContacts.FindNext() as Outlook.ContactItem;
                        if (item != null)
                            col.Add(item);
                    } while (item != null);
                }
			}
			catch (Exception)
			{
				//TODO: should not get here.
			}

			return col;

		}

		public Group GetGoogleGroupById(string id)
		{
			//return _googleGroups.FindById(new AtomId(id)) as Group;
            foreach (Group group in _googleGroups)
            {
                if (group.GroupEntry.Id.Equals(new AtomId(id)))
                    return group;
            }
            return null;
		}

		public Group GetGoogleGroupByName(string name)
		{
			foreach (Group group in _googleGroups)
			{
				if (group.Title == name)
					return group;
			}
			return null;
		}

        public Contact GetGoogleContactById(string id)
        {
            foreach (Contact contact in _googleContacts)
            {
                if (contact.ContactEntry.Id.Equals(new AtomId(id)))
                    return contact;
            }
            return null;
        }

		public Group CreateGroup(string name)
		{
			Group group = new Group();
			group.Title = name;
			group.GroupEntry.Dirty = true;
			return group;
		}

		public static bool AreEqual(Outlook.ContactItem c1, Outlook.ContactItem c2)
		{
			return c1.Email1Address == c2.Email1Address;
		}
		public static int IndexOf(Collection<Outlook.ContactItem> col, Outlook.ContactItem outlookContact)
		{

			for (int i = 0; i < col.Count; i++)
			{
				if (AreEqual(col[i], outlookContact))
					return i;
			}
			return -1;
		}

		internal void DebugContacts()
		{
			string msg = "DEBUG INFORMATION\nPlease submit to developer:\n\n{0}\n{1}\n{2}";

			string oCount = "Outlook Contact Count: " + _outlookContacts.Count;
			string gCount = "Google Contact Count: " + _googleContacts.Count;
			string mCount = "Matches Count: " + _matches.Count;

			MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}

	internal enum SyncOption
	{
		MergePrompt,
		MergeOutlookWins,
		MergeGoogleWins,
		OutlookToGoogleOnly,
		GoogleToOutlookOnly,
	}
}
