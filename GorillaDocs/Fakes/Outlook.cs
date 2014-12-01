using GorillaDocs.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GorillaDocs.Fakes
{
    public class Outlook : GorillaDocs.Models.Outlook
    {
        public Contact GetContact()
        {
            return new Contact() { FullName = "John Smith", CompanyName = "Acme", Title = "Manager", EmailAddress = "John.smith@email.com", PhoneNumber = "999 999 999", Address = "123 Some St\nOurTown State 2000" };
        }
    }
}
