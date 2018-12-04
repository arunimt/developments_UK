using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;

namespace ConvertDocSetToFolder
{
    class Utility
    {
        public ClientContext SPOConnectToken()
        {
#pragma warning disable CS0618 // Type or member is obsolete
            string url = ConfigurationSettings.AppSettings["SiteURL"];
#pragma warning restore CS0618 // Type or member is obsolete
            ClientContext context = null;

            try
            {
                //get site
                Uri siteUri = new Uri(url);

                //Get the realm for the URL
                string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

                //Get the access token for the URL.  Requires this app to be registered with the tenant
                string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

                //Get client context with access token
                context = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken);

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                Console.Write(ex.InnerException);
            }
            return context;
        }

        public ClientContext SPOConnect()
        {
            var securePassword = new SecureString();
#pragma warning disable CS0618 // Type or member is obsolete
            string username = ConfigurationSettings.AppSettings["User"];
            //string password = ConfigurationSettings.AppSettings["Password"];
            string password = getPassword(username);
            string url = ConfigurationSettings.AppSettings["SiteURL"];
            ClientContext context = null;
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            try
            {
                var onlineCredentials = new SharePointOnlineCredentials(ConfigurationSettings.AppSettings["User"], securePassword);
                context = new ClientContext(ConfigurationSettings.AppSettings["SiteURL"]);
                context.Credentials = onlineCredentials;
#pragma warning restore CS0618 // Type or member is obsolete
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return context;
        }

        public string getPassword(string username)
        {
            string pass = "";
            string str = "Enter your password for " + username + " : ";
            Console.Write(str);
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    Console.Write("\b");
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);

            return pass;
        }
    }
}
