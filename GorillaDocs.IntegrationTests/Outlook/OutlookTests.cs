using GorillaDocs.Word;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Net.NetworkInformation;

namespace GorillaDocs.IntegrationTests
{
    [TestFixture]
    public class OutlookTests
    {
        [Test]
        public void Get_Exchange_User()
        {
            const string path = "LDAP://epa.nsw.gov.au";
            //const string path = "LDAP://macroview.com.au";

            var outlook = new Outlook() { LDAPpath = path };
            var contact = outlook.Resolve("Matthew Fitzmaurice");
        }

        [Test]
        public void Does_Domain_Exist()
        {
            var domainId = "macroview.com.au";
            DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, domainId);

            Domain domain;
            try
            {
                domain = Domain.GetDomain(context);
            }
            catch (Exception e)
            {
                throw;
            }
        }

        [Test]
        public void Get_RootDSE()
        {
            DirectoryEntry rootDSE = new DirectoryEntry("LDAP://RootDSE");
            string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
        }

        [Test]
        public void Get_Principal_Context()
        {
            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                // find a user
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, "SomeUserName");

                if (user != null)
                {
                    // do something here....     
                }
            }
        }

        [Test]
        public void Get_Environment_Domain()
        {
            var domain = Environment.UserDomainName;
        }

        [Test]
        public void Get_AD_Domain()
        {
            var domain = System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain().Name;
        }

        [Test]
        public void Get_IPGlobal_Properties()
        {
            var ipproperties = IPGlobalProperties.GetIPGlobalProperties();
            var DomainName = ipproperties.DomainName;
            var HostName = ipproperties.HostName;
        }

        [Test]
        public void Get_Domain_Name_if_not_connected_to_network()
        {
            //http://stackoverflow.com/questions/508911/machines-domain-name-in-net
            var t = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
        }

    }
}
