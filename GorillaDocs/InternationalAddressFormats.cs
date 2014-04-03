using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace GorillaDocs
{
    public class InternationalAddressFormats
    {
        //http://msdn.microsoft.com/en-us/library/cc195167.aspx

        readonly string Street;
        readonly string City;
        readonly string State_or_Province;
        readonly string PostalCode;
        readonly string Country;

        public InternationalAddressFormats(string STREET_ADDRESS, string LOCALITY, string STATE_OR_PROVINCE, string POSTAL_CODE, string COUNTRY)
        {
            this.Street = STREET_ADDRESS;
            this.City = LOCALITY;
            this.State_or_Province = STATE_OR_PROVINCE;
            this.PostalCode = POSTAL_CODE;
            this.Country = COUNTRY;
        }

        public string GetAddress()
        {
            var cultureInfos = CultureInfo.GetCultures(CultureTypes.AllCultures).Where(c => c.DisplayName.Contains(this.Country));
            foreach (var culture in cultureInfos)
                return GetAddress(culture.Name);
            return GetAddress("en-US"); // Default
        }

        public string GetAddress(string CultureName)
        {
            string value = GetAddressString(CultureName);
            value = value.Trim();
            value = value.Replace("  ", " ");
            return value;
        }

        string GetAddressString(string CultureName)
        {
            switch (CultureName)
            {
                case "fr-FR": //French
                    return string.Format("{0}\n{3} {1} {2}\n{4}", Street, City, State_or_Province, PostalCode, Country);
                case "es-ES": //Spanish - Spain
                case "fr-CH": //French - Switzerland
                case "de-CH": //German - Switzerland
                case "sv-SE": //Swedish
                case "tr-TR": //Turkish
                    return string.Format("{0}\n{3} {1}\n{4}", Street, City, State_or_Province, PostalCode, Country);
                case "ro-RO": //Romanian
                    return string.Format("{0}\n{3} {1}\n{2}\n{4}", Street, City, State_or_Province, PostalCode, Country);
                case "ru-RU": //Russian
                    return string.Format("{4}\n{3}\n{2} {1}\n{0}", Street, City, State_or_Province, PostalCode, Country);
                case "en-AU": //English - Australia
                    return string.Format("{0}\n{1} {2} {3}\n{4}", Street, City, State_or_Province, PostalCode, Country);
                default:
                    return string.Format("{0}\n{1}, {2} {3}\n{4}", Street, City, State_or_Province, PostalCode, Country);
            }
        }
    }
}
