using System;
using System.Collections.Generic;
using System.Text;
using Google.GData.Client;
using Google.GData.Contacts;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
	internal static class ContactSync
	{
        //public static void UpdateContact(ContactEntry source, Outlook.ContactItem destination)
        //{
        //    //// if no email or number, contact will be updated at each sync
        //    //if (source.Emails.Count == 0 && source.Phonenumbers.Count == 0)
        //    //    return;


        //    if (!string.IsNullOrEmpty(source.Title.Text))
        //        destination.FileAs = source.Title.Text;
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
        //public static void UpdateContact(Outlook.ContactItem source, ContactEntry destination)
        //{
        //    //// if no email or number, contact will be updated at each sync
        //    //if (string.IsNullOrEmpty(source.Email1Address) && string.IsNullOrEmpty(source.PrimaryTelephoneNumber))
        //    //    return;

        //    if (source.FileAs != source.Email1Address)
        //        destination.Title.Text = source.FileAs;
        //    else
        //        destination.Title.Text = null;

        //    if (destination.Title.Text == null)
        //        destination.Title.Text = source.FullName;

        //    if (destination.Title.Text == null)
        //        destination.Title.Text = source.CompanyName;

        //    SetEmails(source, destination);

        //    SetPhoneNumbers(source, destination);

        //    SetAddresses(source, destination);

        //    SetCompanies(source, destination);

        //    SetIMs(source, destination);

        //    // CH - Fixed error with invalid xml being sent to google... This may need to be added to everything
        //    destination.Content.Content = String.Format("<![CDATA[{0}]]>", source.Body);
        //}

		public static void SetAddresses(Outlook.ContactItem source, ContactEntry destination)
		{
            destination.PostalAddresses.Clear();

			if (!string.IsNullOrEmpty(source.HomeAddress))
			{
				PostalAddress postalAddress = new PostalAddress();
				postalAddress.Value = source.HomeAddress;
                postalAddress.Primary = destination.PostalAddresses.Count == 0;
				postalAddress.Rel = ContactsRelationships.IsHome;
				destination.PostalAddresses.Add(postalAddress);
			}

			if (!string.IsNullOrEmpty(source.BusinessAddress))
			{
				PostalAddress postalAddress = new PostalAddress();
				postalAddress.Value = source.BusinessAddress;
				postalAddress.Primary = destination.PostalAddresses.Count == 0;
				postalAddress.Rel = ContactsRelationships.IsWork;
				destination.PostalAddresses.Add(postalAddress);
			}

			if (!string.IsNullOrEmpty(source.OtherAddress))
			{
				PostalAddress postalAddress = new PostalAddress();
				postalAddress.Value = source.OtherAddress;
				postalAddress.Primary = destination.PostalAddresses.Count == 0;
				postalAddress.Rel = ContactsRelationships.IsOther;
				destination.PostalAddresses.Add(postalAddress);
			}
		}

		public static void SetIMs(Outlook.ContactItem source, ContactEntry destination)
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

		public static void SetEmails(Outlook.ContactItem source, ContactEntry destination)
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

		public static void SetPhoneNumbers(Outlook.ContactItem source, ContactEntry destination)
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

		public static void SetCompanies(Outlook.ContactItem source, ContactEntry destination)
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

		public static void SetPostalAddress(PostalAddress address, Outlook.ContactItem destination)
		{
			if (address.Rel == ContactsRelationships.IsHome)
			{
				destination.HomeAddress = address.Value;
				if (address.Primary)
					destination.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olHome;
				return;
			}
			if (address.Rel == ContactsRelationships.IsWork)
			{
				destination.BusinessAddress = address.Value;
				if (address.Primary)
					destination.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olBusiness;
				return;
			}
			if (address.Rel == ContactsRelationships.IsOther)
			{
				destination.OtherAddress = address.Value;
				if (address.Primary)
					destination.SelectedMailingAddress = Microsoft.Office.Interop.Outlook.OlMailingAddress.olOther;
				return;
			}
		}

	public static void MergeContacts(Outlook.ContactItem master, ContactEntry slave)
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
				slave.Title.Text = master.FileAs;
			else
				slave.Title.Text = null;

			if (slave.Title.Text == null)
				slave.Title.Text = master.FullName;

			if (slave.Title.Text == null)
				slave.Title.Text = master.CompanyName;

			SetEmails(master, slave);

			SetAddresses(master, slave);
			
			SetPhoneNumbers(master, slave);
			
			SetCompanies(master, slave);

			SetIMs(master, slave);

            // CH - Fixed error with invalid xml being sent to google... This may need to be added to everything
            slave.Content.Content = String.Format("<![CDATA[{0}]]>", master.Body);
		}

		public static void MergeContacts(ContactEntry master, Outlook.ContactItem slave)
		{
			//// if no email or number, contact will be updated at each sync
			//if (master.Emails.Count == 0 && master.Phonenumbers.Count == 0)
			//    return;

			if (!string.IsNullOrEmpty(master.Title.Text))
			{
				slave.FileAs = master.Title.Text;
				slave.FullName = master.Title.Text;
			}
			else if (master.Emails.Count > 0)
			{
				slave.FileAs = master.Emails[0].Address;
			}
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

			SetEmails(master, slave);

            //First delete the destination phone numbers having secondary phone numbers
            slave.HomeTelephoneNumber = null;     //secondary: destination.Home2TelephoneNumber
            slave.BusinessTelephoneNumber = null; //secondary: destination.Business2TelephoneNumber

			foreach (PhoneNumber phone in master.Phonenumbers)
			{                
				SetPhoneNumber(phone, slave);
			}

			foreach (PostalAddress address in master.PostalAddresses)
			{
				SetPostalAddress(address, slave);
			}

			slave.Companies = string.Empty;
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

            //ToDo: Sync some fields from second Outlook contact tab, e.g. Birthday, Anniversary, NickName, Department, Partner, Web Address URL, Office, ...

			slave.Body = master.Content.Content;
		}

		public static void SetEmails(ContactEntry source, Outlook.ContactItem destination)
		{
			if (source.Emails.Count > 0)
			{
				destination.Email1Address = source.Emails[0].Address;
				destination.Email1DisplayName = source.Emails[0].Label;
			}
            else
            {
                destination.Email1Address = string.Empty;
                destination.Email1DisplayName = string.Empty;
            }

			if (source.Emails.Count > 1)
            {
				destination.Email2Address = source.Emails[1].Address;
                destination.Email2DisplayName = source.Emails[1].Label;
            }
            else
            {
                destination.Email2Address = string.Empty;
                destination.Email2DisplayName = string.Empty;
            }
			if (source.Emails.Count > 2)
            {
				destination.Email3Address = source.Emails[2].Address;
                destination.Email3DisplayName = source.Emails[2].Label;
            }
            else
            {
                destination.Email3Address = string.Empty;
                destination.Email3DisplayName = string.Empty;
            }
		}

	}
}
