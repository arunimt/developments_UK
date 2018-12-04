using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace GetDocumentCount
{
    class SPOUtility
    {
        public ClientContext SPOConnectToken(Log_Helper log, string url)
        {
            
            ClientContext context = null;
            log.Info("Connecting to SharePoint Online");

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

                log.Info("Connection to SharePoint Online created successfully URL: " + url);
            }
            catch (Exception ex)
            {
                log.Error("Connection to SharePoint Online failed URL: " + url);
                log.Error(ex.Message);
            }
            return context;
        }


        public ClientContext SPOConnect(Log_Helper log, string url, string username, string password)
        {
            var securePassword = new SecureString();
            ClientContext context = null;
            log.Info("Login into SharePoint Online using " + username);
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            securePassword.MakeReadOnly();
            log.Info("Validating Credential");
            try
            {
                var onlineCredentials = new SharePointOnlineCredentials(username, securePassword);
                var cookie = onlineCredentials.GetAuthenticationCookie(new Uri(url));
                //var Credentials = new SharePointOnlineCredentials(username, password);
                context = new ClientContext(url)
                {
                    Credentials = onlineCredentials
                };
                log.Info("Connection to SharePoint Online created successfully URL: " + url);
            }
            catch (Exception ex)
            {
                log.Error("Connection to SharePoint Online failed URL: " + url);
                log.Error(ex.Message);
            }
            return context;
        }
    }
}
