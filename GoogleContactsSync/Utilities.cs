using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Net;
using System.IO;
using Google.GData.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Collections.ObjectModel;

namespace WebGear.GoogleContactsSync
{
    internal static class Utilities
    {
        private static string tempPhotoPath = AppDomain.CurrentDomain.BaseDirectory + "\\TempOutlookContactPhoto.jpg";

        public static byte[] BitmapToBytes(Bitmap bitmap)
        {
            //bitmap
            MemoryStream stream = new MemoryStream();
            bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
            return stream.ToArray();
        }

        public static bool HasPhoto(ContactEntry googleContact)
        {
            if (googleContact.PhotoUri == null)
                return false;
            return true;
        }
        public static bool HasPhoto(Outlook.ContactItem outlookContact)
        {
            return outlookContact.HasPicture;
        }

        public static bool SaveGooglePhoto(Syncronizer sync, ContactEntry googleContact, Image image)
        {
            if (googleContact.PhotoEditUri == null)
                throw new Exception("Must reload contact from google.");

            try
            {
                WebClient client = new WebClient();
                client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.AuthToken);
                client.Headers.Add(HttpRequestHeader.ContentType, "image/*");
                Bitmap pic = new Bitmap(image);
                Stream s = client.OpenWrite(googleContact.PhotoEditUri.AbsoluteUri, "PUT");
                byte[] bytes = BitmapToBytes(pic);
                s.Write(bytes, 0, bytes.Length);
                s.Flush();
                s.Close();
                s.Dispose();
                client.Dispose();
                pic.Dispose();
            }
            catch
            {
                return false;
            }
            return true;
        }
        public static Image GetGooglePhoto(Syncronizer sync, ContactEntry googleContact)
        {
            if (!HasPhoto(googleContact))
                return null;

            try
            {
                WebClient client = new WebClient();
                client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.AuthToken);
                Stream stream = client.OpenRead(googleContact.PhotoUri.AbsoluteUri);
                BinaryReader reader = new BinaryReader(stream);
                Image image = Image.FromStream(stream);
                reader.Close();
                stream.Close();
                stream.Dispose();
                client.Dispose();

                return image;
            }
            catch
            {
                return null;
            }
        }

        public static bool SetOutlookPhoto(Outlook.ContactItem outlookContact, string fullImagePath)
        {
            try
            {
                outlookContact.AddPicture(fullImagePath);
                outlookContact.Save();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public static bool SetOutlookPhoto(Outlook.ContactItem outlookContact, Image image)
        {
            try
            {
                image.Save(tempPhotoPath);
                return SetOutlookPhoto(outlookContact, tempPhotoPath);
            }
            catch
            {
                return false;
            }
        }
        public static Image GetOutlookPhoto(Outlook.ContactItem outlookContact)
        {
            if (!HasPhoto(outlookContact))
                return null;

            try
            {
                foreach (Outlook.Attachment a in outlookContact.Attachments)
                {
                    // CH Fixed this to Contains, due to outlook picture that looks like "ContactPicture_138382.jpg"
                    if( a.DisplayName.ToUpper().Contains( "CONTACTPICTURE")) 
                    {
                        a.SaveAsFile(tempPhotoPath);
                        using (Image img = Image.FromFile(tempPhotoPath))
                        {
                            return new Bitmap(img);
                        }
                    }
                }
                return null;
            }
            catch
            {
                // There's an error here... If Outlook says it has a contact photo, and we can't get it, Something's broken.

                return null;
            }
        }

        public static Image CropImageGoogleFormat(Image original)
        {
            // crop image to a square in the center

            int width, height, diff;
            Point p;
            Rectangle r;

            if (original.Height == original.Width)
                return original;
            if (original.Height > original.Width)
            {
                // tall image
                width = original.Width;
                height = width;

                diff = original.Height - height;
                p = new Point(0, diff / 2);
                r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
            else
            {
                // flat image
                height = original.Height;
                width = height;

                diff = original.Width - width;
                p = new Point(diff / 2, 0);
                r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
        }
        public static Image CropImage(Image original, Rectangle cropArea)
        {
            Bitmap bmpImage = new Bitmap(original);
            Bitmap bmpCrop = bmpImage.Clone(cropArea, bmpImage.PixelFormat);
            return (Image)(bmpCrop);
        }

        public static void DeleteTempPhoto()
        {
            try
            {
                if (File.Exists(tempPhotoPath))
                    File.Delete(tempPhotoPath);
            }
            catch { }
        }

        public static bool ContainsGroup(Syncronizer sync, ContactEntry googleContact, string groupName)
        {
            GroupEntry groupEntry = sync.GetGoogleGroupByName(groupName);
            if (groupEntry == null)
                return false;
            return ContainsGroup(googleContact, groupEntry);
        }
        public static bool ContainsGroup(ContactEntry googleContact, GroupEntry groupEntry)
        {
            foreach (GroupMembership m in googleContact.GroupMembership)
            {
                if (m.HRef == groupEntry.Id.AbsoluteUri)
                    return true;
            }
            return false;
        }
        public static bool ContainsGroup(Outlook.ContactItem outlookContact, string group)
        {
            if (outlookContact.Categories == null)
                return false;

            return outlookContact.Categories.Contains(group);
        }

        public static Collection<GroupEntry> GetGoogleGroups(Syncronizer sync, ContactEntry googleContact)
        {
            int c = googleContact.GroupMembership.Count;
            Collection<GroupEntry> groups = new Collection<GroupEntry>();
            string id;
            GroupEntry group;
            for (int i = 0; i < c; i++)
            {
                id = googleContact.GroupMembership[i].HRef;
                group = sync.GetGoogleGroupById(id);

                groups.Add(group);
            }
            return groups;
        }
        public static void AddGoogleGroup(ContactEntry googleContact, GroupEntry groupEntry)
        {
            if (ContainsGroup(googleContact, groupEntry))
                return;

            GroupMembership m = new GroupMembership();
            m.HRef = groupEntry.Id.AbsoluteUri;
            googleContact.GroupMembership.Add(m);
        }
        public static void RemoveGoogleGroup(ContactEntry googleContact, GroupEntry groupEntry)
        {
            if (!ContainsGroup(googleContact, groupEntry))
                return;

            // TODO: broken. removes group membership but does not remove contact
            // from group in the end.

            // look for id
            GroupMembership mem;
            for (int i = 0; i < googleContact.GroupMembership.Count; i++)
            {
                mem = googleContact.GroupMembership[i];
                if (mem.HRef == groupEntry.Id.AbsoluteUri)
                {
                    googleContact.GroupMembership.Remove(mem);
                    return;
                }
            }
            throw new Exception("Did not find group");
        }

        public static string[] GetOutlookGroups(Outlook.ContactItem outlookContact)
        {
            if (outlookContact.Categories == null)
                return new string[] { };

            string[] categories = outlookContact.Categories.Split(',');
            for (int i = 0; i < categories.Length; i++)
            {
                categories[i] = categories[i].Trim();
            }
            return categories;
        }
        public static void AddOutlookGroup(Outlook.ContactItem outlookContact, string group)
        {
            if (ContainsGroup(outlookContact, group))
                return;

            // append
            if (outlookContact.Categories == null)
                outlookContact.Categories = "";
            if (outlookContact.Categories != "")
                outlookContact.Categories += ", " + group;
            else
                outlookContact.Categories += group;
        }
        public static void RemoveOutlookGroup(Outlook.ContactItem outlookContact, string group)
        {
            if (!ContainsGroup(outlookContact, group))
                return;

            outlookContact.Categories = outlookContact.Categories.Replace(", " + group, "");
            outlookContact.Categories = outlookContact.Categories.Replace(group, "");
        }
    }
}
