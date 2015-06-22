using GorillaDocs.Models;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Runtime.InteropServices;
using Ol = Microsoft.Office.Interop.Outlook;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class Outlook : GorillaDocs.Models.Outlook
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetActiveWindow();
        [DllImport("user32.dll")]
        static extern IntPtr SetActiveWindow(IntPtr handle);

        const string PR_INITIALS = "0x3A0A001F";
        const string PR_SMTP_ADDRESS = "0x39FE001E";
        const string PR_PRIMARY_FAX_NUMBER = "0x3A23001F";
        const string PR_COUNTRY = "0x3A26001F";

        public Outlook() { }
        public Wd.Application WordApplication { get; set; }
        public string LDAPpath { get; set; }

        public Contact GetContact()
        {
            try
            {
                IntPtr handle = GetActiveWindow();
                string value = WordApplication.GetAddress(Type.Missing, "<PR_EMAIL_ADDRESS>");
                SetActiveWindow(handle);

                if (string.IsNullOrEmpty(value))
                    return null;
                return Resolve(value);
            }
            catch (COMException ex)
            {
                if (ex.Message == "Retrieving the COM class factory for component with CLSID {0006F03A-0000-0000-C000-000000000046} failed due to the following error: 80080005.")
                    throw new COMException("Permission error. This error is most likely due to either Microsoft Word or Microsoft Outlook running as Administrator. Both applications must be running in the same context.");
                throw;
            }
        }

        public Contact Resolve(string fullname)
        {
            var app = GetApp();
            if (string.IsNullOrEmpty(app.Name)) throw new InvalidOperationException("Unable to create Outlook session.");

            try
            {
                OutlookSecurityManager.securityManager.ConnectTo(app);
                OutlookSecurityManager.securityManager.DisableOOMWarnings = true;
            }
            catch (Exception)
            {
                System.Threading.Thread.Sleep(1000); // Wait and Try again..
                OutlookSecurityManager.securityManager.ConnectTo(app);
                OutlookSecurityManager.securityManager.DisableOOMWarnings = true;
            }
            try
            {
                OutlookSecurityManager.securityManager.ConnectTo(app);
                OutlookSecurityManager.securityManager.DisableOOMWarnings = true;
                return SecureResolve(app, fullname, LDAPpath);
            }
            finally
            {
                OutlookSecurityManager.securityManager.DisableOOMWarnings = false;
                OutlookSecurityManager.securityManager.Disconnect(app);
            }
        }

        static Ol.Application GetApp()
        {
            try
            {
                return new Ol.Application();
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == -2146959355)
                    throw new InvalidOperationException("Are you debugging as the local Administrator? If so, Outlook must be running with the same credentials");
                // If this line errors with COMException 80080005, ensure that Outlook is started with 'Run as Administrator' http://stackoverflow.com/questions/6369689/outlook-comexception
                throw;
            }
        }
        static Contact SecureResolve(Ol.Application app, string fullname, string LDAPpath)
        {
            var ns = app.GetNamespace("MAPI");
            var recipient = ns.CreateRecipient(fullname);
            recipient.Resolve();
            if (!recipient.Resolved)
                throw new InvalidOperationException(string.Format("Unable to resolve '{0}'", fullname));
            return CreateContact(recipient, LDAPpath);
        }

        static Contact CreateContact(Ol.Recipient recipient, string LDAPpath)
        {
            if (recipient.DisplayType != Ol.OlDisplayType.olUser && recipient.DisplayType != Ol.OlDisplayType.olRemoteUser)
                throw new InvalidOperationException("The recipient must be an individual.");

            var OutlookContact = recipient.AddressEntry.GetContact(); // Does not include GAL contacts.
            if (OutlookContact == null)
                return CreateGALContact(recipient, LDAPpath);
            else
                return CreateOutlookContact(OutlookContact);
        }

        static Contact CreateOutlookContact(Ol.ContactItem item)
        {
            var contact = new Contact();
            //contact.Office = this.office;
            contact.Title = item.Title;
            contact.Initials = item.Initials;
            contact.FullName = item.FullName;
            if (contact.FullName.Contains(','))
                contact.FullName = string.Format("{0} {1}", item.FirstName, item.LastName);
            contact.FirstName = item.FirstName;
            contact.LastName = item.LastName;
            contact.Position = item.JobTitle;
            contact.CompanyName = item.CompanyName;
            contact.PhoneNumber = GetOutlookPhoneNumber(item);
            contact.FaxNumber = string.IsNullOrEmpty(item.BusinessFaxNumber) ? item.HomeFaxNumber : item.BusinessFaxNumber;
            contact.EmailAddress = GetOutlookEmail(item);
            contact.Address = string.IsNullOrEmpty(item.BusinessAddress) ? item.HomeAddress : item.BusinessAddress;
            contact.StreetAddress1 = string.IsNullOrEmpty(item.BusinessAddressStreet) ? item.MailingAddressStreet : item.BusinessAddressStreet;
            contact.StreetCity = string.IsNullOrEmpty(item.BusinessAddressCity) ? item.MailingAddressCity : item.BusinessAddressCity;
            contact.StreetState = string.IsNullOrEmpty(item.BusinessAddressState) ? item.MailingAddressState : item.BusinessAddressState;
            contact.StreetPostalCode = string.IsNullOrEmpty(item.BusinessAddressPostalCode) ? item.MailingAddressPostalCode : item.BusinessAddressPostalCode;
            contact.StreetCountry = string.IsNullOrEmpty(item.BusinessAddressCountry) ? item.MailingAddressCountry : item.BusinessAddressCountry;
            contact.PostalAddress1 = string.IsNullOrEmpty(item.BusinessAddressPostOfficeBox) ? item.MailingAddressPostOfficeBox : item.BusinessAddressPostOfficeBox;
            contact.PostalCity = string.IsNullOrEmpty(item.BusinessAddressCity) ? item.MailingAddressCity : item.BusinessAddressCity;
            contact.PostalState = string.IsNullOrEmpty(item.BusinessAddressState) ? item.MailingAddressState : item.BusinessAddressState;
            contact.PostalPostalCode = string.IsNullOrEmpty(item.BusinessAddressPostalCode) ? item.MailingAddressPostalCode : item.BusinessAddressPostalCode;
            contact.PostalCountry = string.IsNullOrEmpty(item.BusinessAddressCountry) ? item.MailingAddressCountry : item.BusinessAddressCountry;
            contact.Country = string.IsNullOrEmpty(item.BusinessAddressCountry) ? item.MailingAddressCountry : item.BusinessAddressCountry;

            return contact;
        }

        static Contact CreateGALContact(Ol.Recipient recipient, string LDAPpath)
        {
            var contact = new Contact();
            //contact.Office = this.office;
            var user = recipient.AddressEntry.GetExchangeUser();
            if (user == null)
                throw new NullReferenceException("Unable to find address entry in Exchange for user " + recipient.Name);
            //contact.Title = user.;
            contact.Initials = GetGALProperty(user, PR_INITIALS);
            //contact.FullName = user.Name;
            //if (contact.FullName.Contains(',')) // It's possible that Fullname is set up incorrectly..
            contact.FullName = string.Format("{0} {1}", user.FirstName, user.LastName);
            contact.FirstName = user.FirstName;
            contact.LastName = user.LastName;
            contact.Position = user.JobTitle;
            contact.CompanyName = user.CompanyName;
            contact.PhoneNumber = user.BusinessTelephoneNumber;
            contact.FaxNumber = GetGALProperty(user, PR_PRIMARY_FAX_NUMBER);
            contact.EmailAddress = user.PrimarySmtpAddress;
            contact.Address = GetGALAddress(user);
            contact.StreetAddress1 = user.StreetAddress;
            contact.StreetCity = user.City;
            contact.StreetState = user.StateOrProvince;
            contact.StreetPostalCode = user.PostalCode;
            contact.StreetCountry = GetGALProperty(user, PR_COUNTRY);
            contact.PostalAddress1 = user.StreetAddress;
            contact.PostalCity = user.City;
            contact.PostalState = user.StateOrProvince;
            contact.PostalPostalCode = user.PostalCode;
            contact.PostalCountry = GetGALProperty(user, PR_COUNTRY);
            contact.Country = GetGALProperty(user, PR_COUNTRY);
            contact.Assistant = user.AssistantName;

            PopulateExtensionAttributes(recipient, contact, LDAPpath);
            //contact.Assistant = GetGALProperty(user, PR_ASSISTANT);
            //contact.ExtensionAttribute3 = GetGALProperty(user, PR_ExtensionAttribute3);
            //contact.ExtensionAttribute4 = GetGALProperty(user, PR_ExtensionAttribute4);
            return contact;
        }

        static string GetOutlookPhoneNumber(Ol.ContactItem item)
        {
            string value = item.BusinessTelephoneNumber;
            if (string.IsNullOrEmpty(value))
                value = item.HomeTelephoneNumber;
            if (string.IsNullOrEmpty(value))
                value = item.Business2TelephoneNumber;
            if (string.IsNullOrEmpty(value))
                value = item.Home2TelephoneNumber;
            return value;
        }

        static string GetOutlookEmail(Ol.ContactItem item)
        {
            string value = item.Email1Address;
            if (string.IsNullOrEmpty(value))
                value = item.Email2Address;
            if (string.IsNullOrEmpty(value))
                value = item.Email3Address;
            return value;
        }

        static string GetGALAddress(Microsoft.Office.Interop.Outlook.ExchangeUser user)
        {
            var address = new InternationalAddressFormats(user.StreetAddress, user.City, user.StateOrProvince, user.PostalCode, GetGALProperty(user, PR_COUNTRY));
            string value = address.GetAddress();
            return value;
        }

        [System.Diagnostics.DebuggerStepThrough]
        static string GetGALProperty(Microsoft.Office.Interop.Outlook.ExchangeUser user, string property)
        {
            try
            {
                if (property.Contains("schemas"))
                    return user.PropertyAccessor.GetProperty(property) as string;
                else
                    return user.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/" + property) as string;
            }
            catch
            {
                return string.Empty;
            }
        }

        static void PopulateExtensionAttributes(Ol.Recipient recipient, Contact contact, string LDAPpath)
        {
            try
            {
                string legacyExchangeDN = recipient.Address;
                if (!string.IsNullOrEmpty(LDAPpath))
                    using (var dir = new DirectoryEntry(LDAPpath))
                    {
                        dir.RefreshCache();
                        var adSearch = new DirectorySearcher(dir)
                        {
                            Filter = string.Format("(&(objectClass=user)(legacyExchangeDN={0}))", legacyExchangeDN)
                        };
                        SearchResult result = adSearch.FindOne();
                        contact.ExtensionAttribute3 = GetActiveDirectoryProperty(result, "extensionAttribute3").Trim((char)10, (char)13);
                        contact.ExtensionAttribute4 = GetActiveDirectoryProperty(result, "extensionAttribute4").Trim((char)10, (char)13);
                    }
            }
            catch (Exception ex)
            {
                Message.LogError(ex);
            }
        }

        static string GetActiveDirectoryProperty(SearchResult result, string value)
        {
            ResultPropertyValueCollection resultProperties = result.Properties[value];
            if (resultProperties.Count == 1)
                return resultProperties[0] as string;
            else
                return string.Empty;
        }
    }
}
