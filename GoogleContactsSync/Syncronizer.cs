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

namespace WebGear.GoogleContactsSync
{
	internal class Syncronizer
	{
		public const int OutlookUserPropertyMaxLength = 32;
		public const string OutlookUserPropertyTemplate = "g/con/{0}/";

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


		public delegate void NotificationHandler(string title, string message, EventType eventType);
		public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
		public event NotificationHandler DuplicatesFound;
		public event ErrorNotificationHandler ErrorEncountered;

		private ContactsService _googleService;
		public ContactsService GoogleService
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

		private Collection<Outlook.ContactItem> _outlookContactDuplicates;
		public Collection<Outlook.ContactItem> OutlookContactDuplicates
		{
			get { return _outlookContactDuplicates; }
			set { _outlookContactDuplicates = value; }
		}

		private AtomEntryCollection _googleContacts;
		public AtomEntryCollection GoogleContacts
		{
			get { return _googleContacts; }
		}

		private AtomEntryCollection _googleGroups;
		public AtomEntryCollection GoogleGroups
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

		private ContactMatchList _matches;
		public ContactMatchList Contacts
		{
			get { return _matches; }
		}

		private string _authToken;
		public string AuthToken
		{
			get
			{
				return _authToken;
			}
		}

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
				_googleService = new ContactsService("GoogleContactSyncMod");

			_googleService.setUserCredentials(username, password);
			_authToken = _googleService.QueryAuthenticationToken();

			int maxUserIdLength = Syncronizer.OutlookUserPropertyMaxLength - (Syncronizer.OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
			string userId = _googleService.Credentials.Username;
			if (userId.Length > maxUserIdLength)
				userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.

			_propertyPrefix = string.Format(Syncronizer.OutlookUserPropertyTemplate, userId);
		}

		public void LoginToOutlook()
		{
			Logger.Log("Connecting to Outlook...", EventType.Information);

			try
			{
				if (_outlookApp == null)
				{
					_outlookApp = new Outlook.Application();

					_outlookNamespace = _outlookApp.GetNamespace("mapi");
				}
				/// TODO: RedemptioN?
				_outlookNamespace.Logon("Outlook", null, true, false);
			}
			catch (System.Runtime.InteropServices.COMException)
			{
				try
				{
					// If outlook was closed/terminated inbetween, we will receive an Exception
					// System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
					// so recreate outlook instance
					Logger.Log("Cannot connect to Outlook, creating new instance....", EventType.Information);
					_outlookApp = new Outlook.Application();
					_outlookNamespace = _outlookApp.GetNamespace("mapi");
					_outlookNamespace.Logon();
				}
				catch (Exception ex)
				{
					string message = String.Format("Cannot connect to Outlook: {0}.\nPlease restart GO Contact Sync and try again. If error persists, please inform developers on SourceForge.", ex.Message);
					// Error again? We need full stacktrace, display it!
					ErrorHandler.Handle(new ApplicationException(message, ex));
				}
			}
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
		}

		public void LoadOutlookContacts()
		{
			Logger.Log("Loading Outlook contacts...", EventType.Information);
			Outlook.MAPIFolder contactsFolder = _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
			_outlookContacts = contactsFolder.Items;

			//FilterOutlookContactDuplicates();
		}
		/// <summary>
		/// Moves duplicates from _outlookContacts to _outlookContactDuplicates
		/// </summary>
		private void FilterOutlookContactDuplicates()
		{
			_outlookContactDuplicates = new Collection<Outlook.ContactItem>();

			if (_outlookContacts.Count < 2)
				return;

			Outlook.ContactItem main, other;
			bool found = true;
			int index = 0;

			while (found)
			{
				found = false;

				for (int i = index; i <= _outlookContacts.Count - 1; i++)
				{
					main = _outlookContacts[i] as Outlook.ContactItem;

					// only look forward
					for (int j = i + 1; j <= _outlookContacts.Count; j++)
					{
						other = _outlookContacts[j] as Outlook.ContactItem;

						if (other.FileAs == main.FileAs &&
							other.Email1Address == main.Email1Address)
						{
							_outlookContactDuplicates.Add(other);
							_outlookContacts.Remove(j);
							found = true;
							index = i;
							break;
						}
					}
					if (found)
						break;
				}
			}
		}

		public void LoadGoogleContacts()
		{
			try
			{
				Logger.Log("Loading Google Contacts...", EventType.Information);
				ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
				query.NumberToRetrieve = 256;
				query.StartIndex = 0;
				query.ShowDeleted = false;
				//query.OrderBy = "lastmodified";

				ContactsFeed feed;
				feed = _googleService.Query(query);
				_googleContacts = feed.Entries;
				while (feed.Entries.Count == query.NumberToRetrieve)
				{
					query.StartIndex = _googleContacts.Count;
					feed = _googleService.Query(query);
					foreach (AtomEntry a in feed.Entries)
					{
						_googleContacts.Add(a);
					}
				}
			}
			catch (System.Net.WebException ex)
			{
				string message = string.Format("Cannot connect to Google: {0}\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!", ex.Message);
				Logger.Log(message, EventType.Error);
			}
		}
		public void LoadGoogleGroups()
		{
			Logger.Log("Loading Google Groups...", EventType.Information);
			GroupsQuery query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
			query.NumberToRetrieve = 256;
			query.StartIndex = 0;
			query.ShowDeleted = false;

			GroupsFeed feed;
			feed = _googleService.Query(query);
			_googleGroups = feed.Entries;
			while (feed.Entries.Count == query.NumberToRetrieve)
			{
				query.StartIndex = _googleGroups.Count;
				feed = _googleService.Query(query);
				foreach (AtomEntry a in feed.Entries)
				{
					_googleGroups.Add(a);
				}
			}
		}

		public void Load()
		{
			LoadOutlookContacts();
			LoadGoogleContacts();
			LoadGoogleGroups();

			try
			{
				_matches = ContactsMatcher.MatchContacts(this);
			}
			catch (DuplicateDataException ex)
			{
				Logger.Log(ex.Message, EventType.Error);
				if (DuplicatesFound != null)
					DuplicatesFound("Outlook duplicates found", ex.Message, EventType.Error);
			}
		}

		public void Sync()
		{
			_syncedCount = 0;
			_deletedCount = 0;

			Load();

			if (_matches == null)
				return;

			if (_syncProfile.Length == 0)
			{
				Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
				return;
			}

			_totalCount = _matches.Count;

			Logger.Log("Syncing groups...", EventType.Information);
			ContactsMatcher.SyncGroups(this);

			Logger.Log("Syncing contacts...", EventType.Information);
			ContactsMatcher.SyncContacts(this);

			SaveContacts(_matches);
		}

		public void SaveContacts(ContactMatchList contacts)
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
						ErrorEncountered("Error", ex, EventType.Error);
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

				if (match.GoogleContact.IsDirty() || !match.OutlookContact.Saved)
					_syncedCount++;

				if (match.GoogleContact.IsDirty())// || googleChanged)
				{
					//google contact was modified. save.
					SaveGoogleContact(match);
					Logger.Log("Saved Google contact: \"" + match.GoogleContact.Title.Text + "\".", EventType.Information);
				}

				if (!match.OutlookContact.Saved)// || outlookChanged)
				{
					match.OutlookContact.Save();
					ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

					ContactEntry updatedEntry = match.GoogleContact.Update() as ContactEntry;
					match.GoogleContact = updatedEntry;

					ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
					match.OutlookContact.Save();
					Logger.Log("Saved Outlook contact: \"" + match.OutlookContact.FileAs + "\".", EventType.Information);

					//TODO: this will cause the google contact to be updated on next run because Outlook's contact will be marked as saved later that Google's contact.
				}

				// save photos
				SaveContactPhotos(match);
			}
			else if (match.GoogleContact == null && match.OutlookContact != null)
			{
				if (ContactPropertiesUtils.GetOutlookGoogleContactId(this, match.OutlookContact) != null && _syncDelete)
				{
					_deletedCount++;
					string name = match.OutlookContact.FileAs;
					// peer google contact was deleted, delete outlook contact
					match.OutlookContact.Delete();
					Logger.Log("Deleted Outlook contact: \"" + name + "\".", EventType.Information);
				}
			}
			else if (match.GoogleContact != null && match.OutlookContact == null)
			{
				if (ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, match.GoogleContact) != null && _syncDelete)
				{
					_deletedCount++;
					// peer outlook contact was deleted, delete google contact
					match.GoogleContact.Delete();
					Logger.Log("Deleted Google contact: \"" + match.GoogleContact.Title.Text + "\".", EventType.Information);
				}
			}
			else
			{
				//TODO: ignore for now: throw new ArgumentNullException("To save contacts both ContactMatch peers must be present.");
				Logger.Log("Both Google and Outlook contact: \"" + match.GoogleContact.Title.Text + "\" have been changed! Not implemented yet.", EventType.Warning);
			}
		}
		public void SaveGoogleContact(ContactMatch match)
		{
			//check if this contact was not yet inserted on google.
			if (match.GoogleContact.Id.Uri == null)
			{
				//insert contact.
				Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

				try
				{
					ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

					ContactEntry createdEntry = (ContactEntry)_googleService.Insert(feedUri, match.GoogleContact);
					match.GoogleContact = createdEntry;

					ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
					match.OutlookContact.Save();
				}
				catch (Exception ex)
				{
					string xml = GetContactXml(match.GoogleContact);
					string newEx = String.Format("Error saving NEW Google contact: {0}\n{1}", ex.Message, xml);
					throw new ApplicationException(newEx, ex);
				}
			}
			else
			{
				try
				{
					//contact already present in google. just update
					ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

					//TODO: this will fail if original contact had an empty name or rpimary email address.
					ContactEntry updatedEntry = match.GoogleContact.Update() as ContactEntry;
					match.GoogleContact = updatedEntry;

					ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
					match.OutlookContact.Save();
				}
				catch (Exception ex)
				{
					//match.GoogleContact.Summary
					string xml = GetContactXml(match.GoogleContact);
					string newEx = String.Format("Error saving EXISTING Google contact: {0}\n{1}", ex.Message, xml);
					throw new ApplicationException(newEx, ex);
				}
			}
		}

		private string GetContactXml(ContactEntry contactEntry)
		{
			MemoryStream ms = new MemoryStream();
			contactEntry.SaveToXml(ms);
			StreamReader sr = new StreamReader(ms);
			ms.Seek(0, SeekOrigin.Begin);
			string xml = sr.ReadToEnd();
			return xml;
		}

		public void SaveGoogleContact(ContactEntry googleContact)
		{
			//check if this contact was not yet inserted on google.
			if (googleContact.Id.Uri == null)
			{
				//insert contact.
				Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

				try
				{
					ContactEntry createdEntry = (ContactEntry)_googleService.Insert(feedUri, googleContact);
				}
				catch
				{
					//TODO: save google contact xml for diagnistics
					throw;
				}
			}
			else
			{
				try
				{
					//contact already present in google. just update
					//TODO: this will fail if original contact had an empty name or rpimary email address.
					ContactEntry updatedEntry = googleContact.Update() as ContactEntry;
				}
				catch
				{
					//TODO: save google contact xml for diagnistics
					throw;
				}
			}
		}

		public void SaveContactPhotos(ContactMatch match)
		{
			bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
			bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

			if (!hasGooglePhoto && !hasOutlookPhoto)
				return;
			else if (hasGooglePhoto && !hasOutlookPhoto)
			{
				// add google photo to outlook
				Image googlePhoto = Utilities.GetGooglePhoto(this, match.GoogleContact);
				Utilities.SetOutlookPhoto(match.OutlookContact, googlePhoto);
				match.OutlookContact.Save();

				googlePhoto.Dispose();
			}
			else if (!hasGooglePhoto && hasOutlookPhoto)
			{
				// add outlook photo to google
				Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
				if (outlookPhoto != null)
				{
					outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
					bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
					if (!saved)
						throw new Exception("Could not save");

					outlookPhoto.Dispose();
				}
			}
			else
			{
				// TODO: if both contacts have photos and one is updated, the
				// other will not be updated.
			}

			//Utilities.DeleteTempPhoto();
		}

		///// <summary>
		///// Syncs groups. googleChanged and outlookChanged are needed because not always
		///// when a google contact's groups have changed does it become dirty.
		///// </summary>
		///// <param name="match"></param>
		///// <param name="googleChanged"></param>
		///// <param name="outlookChanged"></param>
		//public void SaveContactGroups(ContactMatch match, out bool googleChanged, out bool outlookChanged)
		//{
		//    // get google groups
		//    Collection<GroupEntry> groups = Utilities.GetGoogleGroups(this, match.GoogleContact);

		//    // get outlook categories
		//    string[] cats = Utilities.GetOutlookGroups(match.OutlookContact);

		//    googleChanged = false;
		//    outlookChanged = false;

		//    if (groups.Count == 0 && cats.Length == 0)
		//        return;

		//    switch (SyncOption)
		//    {
		//        case SyncOption.MergeOutlookWins:
		//            //overwrite google contact
		//            OverwriteGoogleContactGroups(match.GoogleContact, groups, cats);
		//            googleChanged = true;
		//            break;
		//        case SyncOption.MergeGoogleWins:
		//            //overwrite outlook contact
		//            OverwriteOutlookContactGroups(match.OutlookContact, cats, groups);
		//            outlookChanged = true;
		//            break;
		//        case SyncOption.GoogleToOutlookOnly:
		//            //overwrite outlook contact
		//            OverwriteOutlookContactGroups(match.OutlookContact, cats, groups);
		//            outlookChanged = true;
		//            break;
		//        case SyncOption.OutlookToGoogleOnly:
		//            //overwrite google contact
		//            OverwriteGoogleContactGroups(match.GoogleContact, groups, cats);
		//            googleChanged = true;
		//            break;
		//        case SyncOption.MergePrompt:
		//            //TODO: we can not rely on previously chosen option as it may be different for each contact.
		//            if (CResolution == null)
		//            {
		//                //promp for sync option
		//                ConflictResolver r = new ConflictResolver();
		//                CResolution = r.Resolve(match.OutlookContact, match.GoogleContact);
		//            }
		//            switch (CResolution)
		//            {
		//                case ConflictResolution.Cancel:
		//                    break;
		//                case ConflictResolution.OutlookWins:
		//                    //overwrite google contact
		//                    OverwriteGoogleContactGroups(match.GoogleContact, groups, cats);
		//                    //TODO: we don't actually know if groups has changed. if all groups were the same it hasn't changed
		//                    googleChanged = true;
		//                    break;
		//                case ConflictResolution.GoogleWins:
		//                    //overwrite outlook contact
		//                    OverwriteOutlookContactGroups(match.OutlookContact, cats, groups);
		//                    //TODO: we don't actually know if groups has changed. if all groups were the same it hasn't changed
		//                    outlookChanged = true;
		//                    break;
		//                default:
		//                    break;
		//            }
		//            break;
		//    }
		//}

		public GroupEntry SaveGoogleGroup(GroupEntry group)
		{
			//check if this group was not yet inserted on google.
			if (group.Id.Uri == null)
			{
				//insert group.
				Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));

				try
				{
					GroupEntry createdEntry = _googleService.Insert(feedUri, group) as GroupEntry;
					return createdEntry;
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
					GroupEntry updatedEntry = group.Update() as GroupEntry;
					return updatedEntry;
				}
				catch
				{
					//TODO: save google group xml for diagnistics
					throw;
				}
			}
		}

		///// <summary>
		///// Updates Google contact's groups
		///// </summary>
		///// <param name="googleContact"></param>
		///// <param name="currentGroups"></param>
		///// <param name="newGroups"></param>
		//public void OverwriteGoogleContactGroups(ContactEntry googleContact, Collection<GroupEntry> currentGroups, string[] newGroups)
		//{
		//    // remove obsolete groups
		//    Collection<GroupEntry> remove = new Collection<GroupEntry>();
		//    bool found;
		//    foreach (GroupEntry group in currentGroups)
		//    {
		//        found = false;
		//        foreach (string cat in newGroups)
		//        {
		//            if (group.Title.Text == cat)
		//            {
		//                found = true;
		//                break;
		//            }
		//        }
		//        if (!found)
		//            remove.Add(group);
		//    }
		//    while (remove.Count != 0)
		//    {
		//        Utilities.RemoveGoogleGroup(googleContact, remove[0]);
		//        remove.RemoveAt(0);
		//    }

		//    // add new groups
		//    GroupEntry g;
		//    foreach (string cat in newGroups)
		//    {
		//        if (!Utilities.ContainsGroup(this, googleContact, cat))
		//        {
		//            // (create and) add group to contact
		//            g = GetGoogleGroupByName(cat);
		//            if (g == null)
		//            {
		//                g = CreateGroup(cat);
		//                SaveGoogleGroup(g);
		//                LoadGoogleGroups();
		//                g = GetGoogleGroupByName(cat);
		//            }
		//            Utilities.AddGoogleGroup(googleContact, g);
		//        }
		//    }
		//}

		/// <summary>
		/// Updates Google contact's groups
		/// </summary>
		/// <param name="googleContact"></param>
		/// <param name="currentGroups"></param>
		/// <param name="newGroups"></param>
		public void OverwriteContactGroups(Outlook.ContactItem master, ContactEntry slave)
		{
			Collection<GroupEntry> currentGroups = Utilities.GetGoogleGroups(this, slave);

			// get outlook categories
			string[] cats = Utilities.GetOutlookGroups(master);

			// remove obsolete groups
			Collection<GroupEntry> remove = new Collection<GroupEntry>();
			bool found;
			foreach (GroupEntry group in currentGroups)
			{
				found = false;
				foreach (string cat in cats)
				{
					if (group.Title.Text == cat)
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
			GroupEntry g;
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
		}

		///// <summary>
		///// Updates Outlook contact's categories (groups)
		///// </summary>
		///// <param name="outlookContact"></param>
		///// <param name="currentGroups"></param>
		///// <param name="newGroups"></param>
		//public void OverwriteOutlookContactGroups(Outlook.ContactItem outlookContact, string[] currentGroups, Collection<GroupEntry> newGroups)
		//{
		//    // remove obsolete groups
		//    Collection<string> remove = new Collection<string>();
		//    bool found;
		//    foreach (string cat in currentGroups)
		//    {
		//        found = false;
		//        foreach (GroupEntry group in newGroups)
		//        {
		//            if (group.Title.Text == cat)
		//            {
		//                found = true;
		//                break;
		//            }
		//        }
		//        if (!found)
		//            remove.Add(cat);
		//    }
		//    while (remove.Count != 0)
		//    {
		//        Utilities.RemoveOutlookGroup(outlookContact, remove[0]);
		//        remove.RemoveAt(0);
		//    }

		//    // add new groups
		//    foreach (GroupEntry group in newGroups)
		//    {
		//        if (!Utilities.ContainsGroup(outlookContact, group.Title.Text))
		//            Utilities.AddOutlookGroup(outlookContact, group.Title.Text);
		//    }
		//}

		/// <summary>
		/// Updates Outlook contact's categories (groups)
		/// </summary>
		/// <param name="outlookContact"></param>
		/// <param name="currentGroups"></param>
		/// <param name="newGroups"></param>
		public void OverwriteContactGroups(ContactEntry master, Outlook.ContactItem slave)
		{
			Collection<GroupEntry> newGroups = Utilities.GetGoogleGroups(this, master);

			List<string> newCats = new List<string>(newGroups.Count);
			foreach (GroupEntry group in newGroups)
			{
				newCats.Add(group.Title.Text);
			}

			slave.Categories = string.Join(", ", newCats.ToArray());
		}

		/// <summary>
		/// Resets associantions of Outlook contacts with Google contacts via user props
		/// and resets associantions of Google contacts with Outlook contacts via extended properties.
		/// </summary>
		public void ResetMatches()
		{
			Debug.Assert(Contacts != null, "Contacts object is null - this should not happen");

			foreach (ContactMatch match in Contacts)
			{
				ResetMatch(match);
			}

			Contacts.Clear();
		}
		public void ResetMatch(ContactMatch match)
		{
			if (match.GoogleContact != null)
			{
				ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, match.GoogleContact);
				SaveGoogleContact(match.GoogleContact);
			}
			if (match.OutlookContact != null)
			{
				ContactPropertiesUtils.ResetOutlookGoogleContactId(this, match.OutlookContact);
				match.OutlookContact.Save();
			}
		}

		public ContactMatch ContactByProperty(string name, string email)
		{
			foreach (ContactMatch m in Contacts)
			{
				if (m.GoogleContact != null &&
					((m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email) ||
					m.GoogleContact.Title.Text == name))
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
		/// <param name="email"></param>
		/// <returns></returns>
		public Collection<Outlook.ContactItem> OutlookContactByProperty(string name, string email)
		{
			Collection<Outlook.ContactItem> col = new Collection<Outlook.ContactItem>();
			foreach (Outlook.ContactItem outlookContact in OutlookContacts)
			{
				if (outlookContact != null && (
					(outlookContact.Email1Address != null && outlookContact.Email1Address == email) ||
					outlookContact.FileAs == name))
				{
					col.Add(outlookContact);
				}
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
			//TODO: optimise by using OutlookContacts.Find

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
			}
			catch (Exception)
			{
				//TODO: should not get here.
			}

			return col;

			//Collection<Outlook.ContactItem> col = new Collection<Outlook.ContactItem>();
			//foreach (Outlook.ContactItem outlookContact in OutlookContacts)
			//{
			//    try
			//    {
			//        if (!(outlookContact is Outlook.ContactItem))
			//            continue;
			//    }
			//    catch (Exception ex)
			//    {
			//        //this is needed because some contacts throw exceptions
			//        continue;
			//    }

			//    if (outlookContact != null && (
			//        outlookContact.Email1Address != null && outlookContact.Email1Address == email))
			//    {
			//        col.Add(outlookContact);
			//    }
			//}
			//return col;
		}

		public GroupEntry GetGoogleGroupById(string id)
		{
			return _googleGroups.FindById(new AtomId(id)) as GroupEntry;

			//foreach (GroupEntry group in _googleGroups)
			//{
			//    if (group.Id.AbsoluteUri == id)
			//        return group;
			//}
			//return null;
		}

		public GroupEntry GetGoogleGroupByName(string name)
		{
			foreach (GroupEntry group in _googleGroups)
			{
				if (group.Title.Text == name)
					return group;
			}
			return null;
		}
		public GroupEntry CreateGroup(string name)
		{
			GroupEntry group = new GroupEntry();
			group.Title.Text = name;
			group.Dirty = true;
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
