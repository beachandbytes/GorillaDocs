using GorillaDocs.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Ol = Microsoft.Office.Interop.Outlook;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public class Outlook : GorillaDocs.Models.Outlook
    {
        readonly Wd.Application app;

        const string PR_INITIALS = "0x3A0A001F";
        const string PR_SMTP_ADDRESS = "0x39FE001E";
        const string PR_PRIMARY_FAX_NUMBER = "0x3A23001F";
        const string PR_COUNTRY = "0x3A26001F";

        public Outlook(Wd.Application app) { this.app = app; }

        public Contact GetContact()
        {
            try
            {
                string value = app.GetAddress(Type.Missing, "<PR_EMAIL_ADDRESS>");
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

        static Contact Resolve(string fullname)
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
                return SecureResolve(app, fullname);
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
        static Contact SecureResolve(Ol.Application app, string fullname)
        {
            var ns = app.GetNamespace("MAPI");
            var recipient = ns.CreateRecipient(fullname);
            recipient.Resolve();
            if (!recipient.Resolved)
                throw new InvalidOperationException(string.Format("Unable to resolve '{0}'", fullname));
            return CreateContact(recipient);
        }

        static Contact CreateContact(Ol.Recipient recipient)
        {
            if (recipient.DisplayType != Ol.OlDisplayType.olUser)
                throw new InvalidOperationException("The recipient must be an individual.");

            var OutlookContact = recipient.AddressEntry.GetContact(); // Does not include GAL contacts.
            if (OutlookContact == null)
                return CreateGALContact(recipient);
            else
                return CreateOutlookContact(OutlookContact);
        }

        static Contact CreateOutlookContact(Ol.ContactItem item)
        {
            var contact = new Contact();
            //contact.Office = this.office;
            contact.Initials = item.Initials;
            contact.FullName = item.FullName;
            if (contact.FullName.Contains(','))
                contact.FullName = string.Format("{0} {1}", item.FirstName, item.LastName);
            contact.FirstName = item.FirstName;
            contact.LastName = item.LastName;
            contact.Title = item.JobTitle;
            contact.CompanyName = item.CompanyName;
            contact.PhoneNumber = GetOutlookPhoneNumber(item);
            contact.FaxNumber = GetOutlookFaxNumber(item);
            contact.EmailAddress = GetOutlookEmail(item);
            contact.Address = GetOutlookAddress(item);
            //contact.OfficeLocation = item.OfficeLocation;
            contact.Country = GetOutlookCountry(item);
            return contact;
        }

        static Contact CreateGALContact(Ol.Recipient recipient)
        {
            var contact = new Contact();
            //contact.Office = this.office;
            var user = recipient.AddressEntry.GetExchangeUser();
            if (user == null)
                throw new NullReferenceException("Unable to find address entry in Exchange for user " + recipient.Name);
            contact.Initials = GetGALProperty(user, PR_INITIALS);
            contact.FullName = user.Name;
            if (contact.FullName.Contains(','))
                contact.FullName = string.Format("{0} {1}", user.FirstName, user.LastName);
            contact.FirstName = user.FirstName;
            contact.LastName = user.LastName;
            contact.Title = user.JobTitle;
            contact.CompanyName = user.CompanyName;
            contact.PhoneNumber = user.BusinessTelephoneNumber;
            contact.FaxNumber = GetGALProperty(user, PR_PRIMARY_FAX_NUMBER);
            contact.EmailAddress = user.PrimarySmtpAddress;
            contact.Address = GetGALAddress(user);
            contact.Country = GetGALProperty(user, PR_COUNTRY);
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

        static string GetOutlookFaxNumber(Ol.ContactItem item)
        {
            string value = item.BusinessFaxNumber;
            if (string.IsNullOrEmpty(value))
                value = item.HomeFaxNumber;
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

        static string GetOutlookAddress(Ol.ContactItem item)
        {
            string value = item.BusinessAddress;
            if (string.IsNullOrEmpty(value))
                value = item.HomeAddress;
            return value;
        }

        static string GetOutlookCountry(Ol.ContactItem item)
        {
            string value = item.BusinessAddressCountry;
            if (string.IsNullOrEmpty(value))
                value = item.HomeAddressCountry;
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
                return user.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/" + property) as string;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
