using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Client;
using Google.GData.Contacts;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.Contacts;

namespace GoContactSyncMod
{
    
	internal static class ContactSync
	{
        private static DateTime outlookDateNone = new DateTime(4501, 1, 1);
        private const string relSpouse = "spouse";
        private const string relChild = "child";
        private const string relAnniversary = "anniversary";
        private const string relHomePage = "home-page";
    
        //public static void UpdateContact(Contact source, Outlook.ContactItem destination)
        //{
        //    //// if no email or number, contact will be updated at each sync
        //    //if (source.Emails.Count == 0 && source.Phonenumbers.Count == 0)
        //    //    return;


        //    if (!string.IsNullOrEmpty(source.Title))
        //        destination.FileAs = source.Title;
        //    else
        //        destination.FileAs = source.Emails[0].Address;

        //    SetEmails(source, destination);

        //    //First delete the destination phone numbers having secondary phone numbers
        //    destination.HomeTelephoneNumber = null;     //secondary: destination.Home2TelephoneNumber
        //    destination.BusinessTelephoneNumber = null; //secondary: destination.Business2TelephoneNumber

        //    foreach (PhoneNumber phone in source.Phonenumbers)
        //    {
        //        SetPhoneNumber(phone, destination);
        //    }

        //    foreach (PostalAddress address in source.PostalAddresses)
        //    {
        //        SetPostalAddress(address, destination);
        //    }

        //    destination.Companies = string.Empty;
        //    foreach (Organization company in source.Organizations)
        //    {
        //        if (company.Primary)
        //        {
        //            destination.CompanyName = company.Name;
        //            destination.JobTitle = company.Title;
        //        }
        //        if (destination.Companies.Length > 0)
        //            destination.Companies += "; ";
        //        destination.Companies += company.Name;
        //    }

        //    destination.IMAddress = "";
        //    foreach (IMAddress im in source.IMs)
        //    {
        //        if (destination.IMAddress.Length > 0)
        //            destination.IMAddress += "; ";
        //        if (!string.IsNullOrEmpty(im.Protocol))
        //            destination.IMAddress += im.Protocol + ": " + im.Address;
        //        destination.IMAddress += im.Address;
        //    }

        //    destination.Body = source.Content.Content;
        //}

        ///// <summary>
        ///// Replaces all properties of <paramref name="destination"/> from corresponding properties of <paramref name="source"/>
        ///// </summary>
        ///// <param name="source"></param>
        ///// <param name="destination"></param>
        //public static void UpdateContact(Outlook.ContactItem source, Contact destination)
        //{
        //    //// if no email or number, contact will be updated at each sync
        //    //if (string.IsNullOrEmpty(source.Email1Address) && string.IsNullOrEmpty(source.PrimaryTelephoneNumber))
        //    //    return;

        //    if (source.FileAs != source.Email1Address)
        //        destination.Title = source.FileAs;
        //    else
        //        destination.Title = null;

        //    if (destination.Title == null)
        //        destination.Title = source.FullName;

        //    if (destination.Title == null)
        //        destination.Title = source.CompanyName;

        //    SetEmails(source, destination);

        //    SetPhoneNumbers(source, destination);

        //    SetAddresses(source, destination);

        //    SetCompanies(source, destination);

        //    SetIMs(source, destination);

        //    // CH - Fixed error with invalid xml being sent to google... This may need to be added to everything
        //    destination.Content.Content = String.Format("<![CDATA[{0}]]>", source.Body);
        //}

		public static void SetAddresses(Outlook.ContactItem source, Contact destination)
		{
            destination.PostalAddresses.Clear();

			if (!string.IsNullOrEmpty(source.HomeAddress))
			{
				StructuredPostalAddress postalAddress = new StructuredPostalAddress();
                postalAddress.Street = source.HomeAddressStreet;
                postalAddress.City = source.HomeAddressCity;
                postalAddress.Postcode = source.HomeAddressPostalCode;
                postalAddress.Country = source.HomeAddressCountry;
                postalAddress.Primary = destination.PostalAddresses.Count == 0;
				postalAddress.Rel = ContactsRelationships.IsHome;
				destination.PostalAddresses.Add(postalAddress);
			}

			if (!string.IsNullOrEmpty(source.BusinessAddress))
			{
				StructuredPostalAddress postalAddress = new StructuredPostalAddress();
                postalAddress.Street = source.BusinessAddressStreet;
                postalAddress.City = source.BusinessAddressCity;
                postalAddress.Postcode = source.BusinessAddressPostalCode;
                postalAddress.Country = source.BusinessAddressCountry;
				postalAddress.Primary = destination.PostalAddresses.Count == 0;
				postalAddress.Rel = ContactsRelationships.IsWork;
				destination.PostalAddresses.Add(postalAddress);
			}

			if (!string.IsNullOrEmpty(source.OtherAddress))
			{
				StructuredPostalAddress postalAddress = new StructuredPostalAddress();
                postalAddress.Street = source.OtherAddressStreet;
                postalAddress.City = source.OtherAddressCity;
                postalAddress.Postcode = source.OtherAddressPostalCode;
                postalAddress.Country = source.OtherAddressCountry;
				postalAddress.Primary = destination.PostalAddresses.Count == 0;
				postalAddress.Rel = ContactsRelationships.IsOther;
				destination.PostalAddresses.Add(postalAddress);
			}
		}

		public static void SetIMs(Outlook.ContactItem source, Contact destination)
		{
            destination.IMs.Clear();

			if (!string.IsNullOrEmpty(source.IMAddress))
			{
				//IMAddress are expected to be in form of ([Protocol]: [Address]; [Protocol]: [Address])
				string[] imsRaw = source.IMAddress.Split(';');
				foreach (string imRaw in imsRaw)
				{
					string[] imDetails = imRaw.Trim().Split(':');
					IMAddress im = new IMAddress();
					if (imDetails.Length == 1)
						im.Address = imDetails[0].Trim();
					else
					{
						im.Protocol = imDetails[0].Trim();
						im.Address = imDetails[1].Trim();
					}
					im.Primary = destination.IMs.Count == 0;
					im.Rel = ContactsRelationships.IsHome;
					destination.IMs.Add(im);
				}
			}
		}

		public static void SetEmails(Outlook.ContactItem source, Contact destination)
		{
            destination.Emails.Clear();

			if (!string.IsNullOrEmpty(source.Email1Address))
			{
				EMail primaryEmail = new EMail(source.Email1Address);
				primaryEmail.Primary = destination.Emails.Count == 0;
				primaryEmail.Rel = ContactsRelationships.IsWork;
				destination.Emails.Add(primaryEmail);
			}

			if (!string.IsNullOrEmpty(source.Email2Address))
			{
				EMail secondaryEmail = new EMail(source.Email2Address);
				secondaryEmail.Primary = destination.Emails.Count == 0;
				secondaryEmail.Rel = ContactsRelationships.IsHome;
				destination.Emails.Add(secondaryEmail);
			}

			if (!string.IsNullOrEmpty(source.Email3Address))
			{
				EMail secondaryEmail = new EMail(source.Email3Address);
				secondaryEmail.Primary = destination.Emails.Count == 0;
				secondaryEmail.Rel = ContactsRelationships.IsOther;
				destination.Emails.Add(secondaryEmail);
			}
		}

		public static void SetPhoneNumbers(Outlook.ContactItem source, Contact destination)
		{
            destination.Phonenumbers.Clear();

			if (!string.IsNullOrEmpty(source.PrimaryTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.PrimaryTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsMobile;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.MobileTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.MobileTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsMobile;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.HomeTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.HomeTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsHome;
				destination.Phonenumbers.Add(phoneNumber);
			}

            if (!string.IsNullOrEmpty(source.Home2TelephoneNumber))
            {
                PhoneNumber phoneNumber = new PhoneNumber(source.Home2TelephoneNumber);
                phoneNumber.Primary = destination.Phonenumbers.Count == 0;
                phoneNumber.Rel = ContactsRelationships.IsHome;
                destination.Phonenumbers.Add(phoneNumber);
            }

			if (!string.IsNullOrEmpty(source.BusinessTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.BusinessTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsWork;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.Business2TelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.Business2TelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsWork;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.HomeFaxNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.HomeFaxNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsHomeFax;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.BusinessFaxNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.BusinessFaxNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsWorkFax;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.OtherTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.OtherTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsOther;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.RadioTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.RadioTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsMobile;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.PagerNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.PagerNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsPager;
				destination.Phonenumbers.Add(phoneNumber);
			}

			if (!string.IsNullOrEmpty(source.CarTelephoneNumber))
			{
				PhoneNumber phoneNumber = new PhoneNumber(source.CarTelephoneNumber);
				phoneNumber.Primary = destination.Phonenumbers.Count == 0;
				phoneNumber.Rel = ContactsRelationships.IsCar;
				destination.Phonenumbers.Add(phoneNumber);
			}
		}

		public static void SetCompanies(Outlook.ContactItem source, Contact destination)
		{
            destination.Organizations.Clear();

			if (!string.IsNullOrEmpty(source.Companies))
			{
				//Companies are expected to be in form of "[Company]; [Company]".
				string[] companiesRaw = source.Companies.Split(';');
				foreach (string companyRaw in companiesRaw)
				{
					Organization company = new Organization();
                    company.Name = (destination.Organizations.Count == 0) ? source.CompanyName : null;
                    company.Title = (destination.Organizations.Count == 0)?source.JobTitle : null;
					company.Primary = destination.Organizations.Count == 0;
					company.Rel = ContactsRelationships.IsWork;
					destination.Organizations.Add(company);
				}
			}

			if (destination.Organizations.Count == 0 && (!string.IsNullOrEmpty(source.CompanyName) || !string.IsNullOrEmpty(source.JobTitle)))
			{
				Organization company = new Organization();
				company.Name = source.CompanyName;
                company.Title = source.JobTitle;
				company.Primary = true;
				company.Rel = ContactsRelationships.IsWork;
				destination.Organizations.Add(company);
			}
		}

		public static void SetPhoneNumber(PhoneNumber phone, Outlook.ContactItem destination)
		{
            if (phone.Primary)
                destination.PrimaryTelephoneNumber = phone.Value;

            if (phone.Rel == ContactsRelationships.IsHome)
            {
                if (destination.HomeTelephoneNumber == null)
                    destination.HomeTelephoneNumber = phone.Value;
                else
                    destination.Home2TelephoneNumber = phone.Value;
            }
            else if (phone.Rel == ContactsRelationships.IsWork)
            {
                if (destination.BusinessTelephoneNumber == null)
                    destination.BusinessTelephoneNumber = phone.Value;
                else
                    destination.Business2TelephoneNumber = phone.Value;
            }
            else if (phone.Rel == ContactsRelationships.IsMobile)
			{
				destination.MobileTelephoneNumber = phone.Value;
				//destination.PrimaryTelephoneNumber = phone.Value;
			}
			else if (phone.Rel == ContactsRelationships.IsWorkFax)
				destination.BusinessFaxNumber = phone.Value;
			else if (phone.Rel == ContactsRelationships.IsHomeFax)
				destination.HomeFaxNumber = phone.Value;
            else if (phone.Rel == ContactsRelationships.IsPager)
				destination.PagerNumber = phone.Value;
			else if (phone.Rel == ContactsRelationships.IsSatellite)
				destination.RadioTelephoneNumber = phone.Value;
			else if (phone.Rel == ContactsRelationships.IsOther)
				destination.OtherTelephoneNumber = phone.Value;
			else if (phone.Rel == ContactsRelationships.IsCar)
				destination.CarTelephoneNumber = phone.Value;
            //else if (phone.Rel == ContactsRelationships.IsVoip)
            //    destination.Business2TelephoneNumber = phone.Value;
            //else no phone category matches
		}

		public static void SetPostalAddress(StructuredPostalAddress address, Outlook.ContactItem destination)
		{
			if (address.Rel == ContactsRelationships.IsHome)
			{
                destination.HomeAddressStreet=address.Street;
                destination.HomeAddressCity=address.City;
                destination.HomeAddressPostalCode=address.Postcode;
                destination.HomeAddressCountry=address.Country;

				if (address.Primary)
					destination.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olHome;
				return;
			}
			if (address.Rel == ContactsRelationships.IsWork)
			{
                destination.BusinessAddressStreet = address.Street;
                destination.BusinessAddressCity = address.City;
                destination.BusinessAddressPostalCode = address.Postcode;
                destination.BusinessAddressCountry = address.Country;

				if (address.Primary)
					destination.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olBusiness;
				return;
			}
			if (address.Rel == ContactsRelationships.IsOther)
			{
                destination.OtherAddressStreet = address.Street;
                destination.OtherAddressCity = address.City;
                destination.OtherAddressPostalCode = address.Postcode;
                destination.OtherAddressCountry = address.Country;

				if (address.Primary)
					destination.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olOther;
				return;
			}
		}

	public static void MergeContacts(Outlook.ContactItem master, Contact slave)
		{
			//// if no email or number, contact will be updated at each sync
            //if (string.IsNullOrEmpty(master.Email1Address) && string.IsNullOrEmpty(master.PrimaryTelephoneNumber))
            //{
            //    if (slave.Emails.Count > 0)
            //    {
            //        Logger.Log("Outlook Contact '" + master.FullNameAndCompany + "' has neither E-Mail address nor phone number. Setting E-Mail address of Google contact: " + slave.Emails[0].Address, EventType.Warning);
            //        master.Email1Address = slave.Emails[0].Address;
            //    }
            //    else
            //    {
            //        Logger.Log("Outlook Contact '" + master.FullNameAndCompany + "' has neither E-Mail address nor phone number. Cannot merge with Google contac: " + slave.Summary, EventType.Error);
            //        return;
            //    }					
            //}

			//TODO: convert to merge as opposed to replace

            if (master.FileAs != master.Email1Address)
            {
                slave.Title = master.FileAs;
            }
            else
                slave.Title = null;

			if (slave.Title == null)
				slave.Title = master.FullName;

			if (slave.Title == null)
				slave.Title = master.CompanyName;

            Name name = new Name();
            name.FullName = master.FullName;

            name.NamePrefix = master.Title;
            name.GivenName = master.FirstName;
            name.AdditonalName = master.MiddleName;
            name.FamilyName = master.LastName;
            name.NameSuffix = master.Suffix;
            slave.Name = name;

            if (master.Birthday.Equals(outlookDateNone)) //earlier also || master.Birthday.Year < 1900
                slave.ContactEntry.Birthday = null;
            else
                slave.ContactEntry.Birthday = master.Birthday.ToString("yyyy-MM-dd");
            slave.ContactEntry.Nickname = master.NickName;
            slave.Location = master.OfficeLocation;
            //Categories are synced separately in Syncronizer.OverwriteContactGroups: slave.Categories = master.Categories;
            slave.ContactEntry.Initials = master.Initials;
            slave.ContactEntry.Language = master.Language;
            //ToDo: Sync department from second Outlook contact tab

			SetEmails(master, slave);

			SetAddresses(master, slave);
			
			SetPhoneNumbers(master, slave);
			
			SetCompanies(master, slave);

			SetIMs(master, slave);

            //First remove anniversary
            foreach (Event ev in slave.ContactEntry.Events)
            {
                if (ev.Relation != null && ev.Relation.Equals(relAnniversary))
                {
                    slave.ContactEntry.Events.Remove(ev);
                    break;
                }
            }
            //Then add it again if existing
            if (!master.Anniversary.Equals(outlookDateNone)) //earlier also || master.Birthday.Year < 1900
            {
                Event ev = new Event();
                ev.Relation = relAnniversary;
                ev.When = new When();
                ev.When.AllDay = true;
                ev.When.StartTime = master.Anniversary.Date;            
                slave.ContactEntry.Events.Add(ev);
            }

            //First remove spouse and child
            foreach (Relation rel in slave.ContactEntry.Relations)
            {
                if (rel.Rel != null && (rel.Rel.Equals(relSpouse) || rel.Rel.Equals(relChild)))
                {
                    slave.ContactEntry.Relations.Remove(rel);
                    break;
                }
            }
            //Then add spouse again if existing
            if (!string.IsNullOrEmpty(master.Spouse))        
            {
                Relation rel = new Relation();
                rel.Rel = relSpouse;
                rel.Value = master.Spouse;                
                slave.ContactEntry.Relations.Add(rel);
            }
            //Then add children again if existing
            if (!string.IsNullOrEmpty(master.Children))               
            {
                Relation rel = new Relation();
                rel.Rel = relChild;
                rel.Value = master.Children;                
                slave.ContactEntry.Relations.Add(rel);
            }

            slave.ContactEntry.Websites.Clear();
            //Just copy the first URL, because Outlook only has 1
            if (master.WebPage != null)
            {
                Website url = new Website();
                url.Href = master.WebPage;
                url.Rel = relHomePage;
                slave.ContactEntry.Websites.Add(url);
            }

            // CH - Fixed error with invalid xml being sent to google... This may need to be added to everything
            //slave.Content = String.Format("<![CDATA[{0}]]>", master.Body);
            //floriwan: Maybe better to jusst esapce the XML instead of putting it in CDATA, because this causes a CDATA added to all my contacts
            if (master.Body != null)
                slave.Content = String.Format(System.Security.SecurityElement.Escape(master.Body));
            else
                slave.Content = null;
		}

		public static void MergeContacts(Contact master, Outlook.ContactItem slave)
		{
			//// if no email or number, contact will be updated at each sync
			//if (master.Emails.Count == 0 && master.Phonenumbers.Count == 0)
			//    return;

			if (!string.IsNullOrEmpty(master.Title))
				slave.FileAs = master.Title;
			else if (master.Emails.Count > 0)
				slave.FileAs = master.Emails[0].Address;
			else
			{
				if (!String.IsNullOrEmpty(slave.Email1Address))
				{
					Logger.Log("Google Contact '" + master.Summary + "' has neither E-Mail address nor phone number. Setting E-Mail address of Outlook contact: " + slave.Email1Address, EventType.Warning);
					master.Emails.Add(new EMail(slave.Email1Address));
					slave.FileAs = master.Emails[0].Address;
				}
				else
				{
					Logger.Log("Google Contact '" + master.Summary + "' has neither E-Mail address nor phone number. Cannot merge with Outlook contact: " + slave.FullNameAndCompany, EventType.Error);
					return;
				}
			}
            
            slave.FullName = master.Name.FullName;
            slave.Title = master.Name.NamePrefix;
            slave.FirstName = master.Name.GivenName;
            slave.MiddleName = master.Name.AdditonalName;
            slave.LastName = master.Name.FamilyName;
            slave.Suffix = master.Name.NameSuffix;
            DateTime birthday;
            DateTime.TryParse(master.ContactEntry.Birthday, out birthday);

            if (birthday != DateTime.MinValue)
                slave.Birthday = birthday;
            else
                slave.Birthday = outlookDateNone;
            slave.NickName = master.ContactEntry.Nickname;
            slave.OfficeLocation = master.Location;
            //Categories are synced separately in Syncronizer.OverwriteContactGroups: slave.Categories = master.Categories;
            slave.Initials = master.ContactEntry.Initials;
            slave.Language = master.ContactEntry.Language;
            //ToDo: Sync department from second Outlook contact tab
            
			SetEmails(master, slave);

            //First delete the destination phone numbers
            slave.HomeTelephoneNumber = null;
            slave.Home2TelephoneNumber = null;
            slave.BusinessTelephoneNumber = null;
            slave.Business2TelephoneNumber = null;
            slave.MobileTelephoneNumber = null;
            slave.BusinessFaxNumber = null;
            slave.HomeFaxNumber = null;
            slave.PagerNumber = null;
            slave.RadioTelephoneNumber = null;
            slave.OtherTelephoneNumber = null;
            slave.CarTelephoneNumber = null;
            
			foreach (PhoneNumber phone in master.Phonenumbers)
			{                
				SetPhoneNumber(phone, slave);
			}

            //ToDo: What if the OutlookContact only has e.g. HomeAddress or BusinessAddress properties set, without the structured postal address? Normally this should happen
            slave.HomeAddress = null;
            slave.HomeAddressStreet = null;
            slave.HomeAddressCity = null;
            slave.HomeAddressPostalCode = null;
            slave.HomeAddressCountry = null;

            slave.BusinessAddress = null;
            slave.BusinessAddressStreet = null;
            slave.BusinessAddressCity = null;
            slave.BusinessAddressPostalCode = null;
            slave.BusinessAddressCountry = null;

            slave.OtherAddress = null;
            slave.OtherAddressStreet = null;
            slave.OtherAddressCity = null;
            slave.OtherAddressPostalCode = null;
            slave.OtherAddressCountry = null;

            slave.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olNone;
			foreach (StructuredPostalAddress address in master.PostalAddresses)
			{
				SetPostalAddress(address, slave);
			}

			slave.Companies = string.Empty;
            slave.CompanyName = string.Empty;
			foreach (Organization company in master.Organizations)
			{
				if (string.IsNullOrEmpty(company.Name) && string.IsNullOrEmpty(company.Title))
					continue;

				if (company.Primary)
                {
					slave.CompanyName = company.Name;
                    slave.JobTitle = company.Title;
                }
				if (!string.IsNullOrEmpty(slave.Companies))
					slave.Companies += "; ";
				slave.Companies += company.Name;
			}

			slave.IMAddress = "";
			foreach (IMAddress im in master.IMs)
			{
				if (!string.IsNullOrEmpty(slave.IMAddress))
					slave.IMAddress += "; ";
				if (!string.IsNullOrEmpty(im.Protocol))
					slave.IMAddress += im.Protocol + ": " + im.Address;
				slave.IMAddress += im.Address;
			}            

            
            slave.Anniversary = outlookDateNone; //set to empty first
            foreach (Event ev in master.ContactEntry.Events)
            {
                if (ev.Relation != null && ev.Relation.Equals(relAnniversary))
                    slave.Anniversary = ev.When.StartTime.Date;
            }

            slave.Children = null;
            slave.Spouse = null;
            foreach (Relation rel in master.ContactEntry.Relations)
            {
                if (rel.Rel != null && rel.Rel.Equals(relChild))
                    slave.Children = rel.Value;
                if (rel.Rel != null && rel.Rel.Equals(relSpouse))
                    slave.Spouse = rel.Value;
            }

            //Just copy the first URL, because Outlook only has 1
            if (master.ContactEntry.Websites.Count > 0)
                slave.WebPage = master.ContactEntry.Websites[0].Href;

			slave.Body = master.Content;
		}

		public static void SetEmails(Contact source, Outlook.ContactItem destination)
		{
            destination.Email1Address = string.Empty;
            destination.Email1DisplayName = string.Empty;

            destination.Email2Address = string.Empty;
            destination.Email2DisplayName = string.Empty;

            destination.Email3Address = string.Empty;
            destination.Email3DisplayName = string.Empty;

			if (source.Emails.Count > 0)
			{
				destination.Email1Address = source.Emails[0].Address;
				destination.Email1DisplayName = source.Emails[0].Label;
			}            

			if (source.Emails.Count > 1)
            {
				destination.Email2Address = source.Emails[1].Address;
                destination.Email2DisplayName = source.Emails[1].Label;
            }
            
			if (source.Emails.Count > 2)
            {
				destination.Email3Address = source.Emails[2].Address;
                destination.Email3DisplayName = source.Emails[2].Label;
            }
            
		}

	}
}
