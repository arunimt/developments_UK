using System;
using System.Configuration;
using System.IO;
using Microsoft.SharePoint.Client;

namespace MergeDocs
{
    public class Program
    {
        private static void Main(string[] args)
        {
            Uri siteUri = new Uri(ConfigurationManager.AppSettings["SiteURL"]);

            //Get the realm for the URL
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);

            //Get the access token for the URL.  Requires this app to be registered with the tenant
            string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

            //Get client context with access token
            var context = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken);
            var ServerRelativeUrl = @"/sites/CommercialDev1/Commercial/hello/Source";
            var files = context.Web.GetFolderByServerRelativeUrl(ServerRelativeUrl).Files;
            ///// Need to query based on a flag
            context.Load(files);
            context.ExecuteQuery();
            using (MemoryStream streamSrc = new MemoryStream())
            {
                foreach (Microsoft.SharePoint.Client.File file in files)
                {
                    ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                    context.Load(file);
                    context.ExecuteQuery();

                    if (data != null)
                    {
                        data.Value.CopyTo(streamSrc);
                    }
                }

                string url = ConfigurationSettings.AppSettings["SiteURL"];

                try
                {
                    var listName = "hello";
                    var folderName = "Destination";
                    var fileName = "xyz.docx";

                    var list = context.Web.Lists.GetByTitle(listName);
                    context.Load(list.RootFolder);
                    context.ExecuteQuery();

                    var targetFileUrl = String.Format("{0}/{1}/{2}", list.RootFolder.ServerRelativeUrl, folderName, fileName);

                    var fileCreationInformation = new FileCreationInformation();

                    //Assign to content byte[] i.e. documentStream
                    fileCreationInformation.Content = streamSrc.ToArray();

                    //Allow owerwrite of document
                    fileCreationInformation.Overwrite = true;

                    //Upload URL
                    fileCreationInformation.Url = targetFileUrl;//siteURL + documentListURL + documentName;

                    Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);

                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                }
            }
        }
    }
}