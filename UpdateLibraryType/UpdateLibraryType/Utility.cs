using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;

namespace UpdateLibraryType
{
    class Utility
    {
        public ClientContext SPOConnect()
        {
            var securePassword = new SecureString();
            //string password = ConfigurationSettings.AppSettings["Password"];
            //string username = ConfigurationSettings.AppSettings["User"];
            //string url = ConfigurationSettings.AppSettings["SiteURL"];
            ClientContext context = null;
            foreach (char c in ConfigurationSettings.AppSettings["Password"])
            {
                securePassword.AppendChar(c);
            }
            try
            {
                var onlineCredentials = new SharePointOnlineCredentials(ConfigurationSettings.AppSettings["User"], securePassword);
                context = new ClientContext(ConfigurationSettings.AppSettings["SiteURL"]);
                context.Credentials = onlineCredentials;
            }
            catch (Exception ex)
            {
            }
            return context;
        }
    }
}
