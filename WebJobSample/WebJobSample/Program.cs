using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.SharePoint.Client;

namespace WebJobSample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Uri siteUri = new Uri(ConfigurationSettings.AppSettings["SiteURL"]);
                //Get the realm for the URL
                string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
                //Get the access token for the URL.  Requires this app to be registered with the tenant
                string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
                //Get client context with access token
                var context = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken);

                var list = context.Web.Lists.GetByTitle(ConfigurationSettings.AppSettings["Library"]);
                context.Load(list);
                context.Load(list.RootFolder.Folders);
                context.ExecuteQuery();
                ListItemCollection items = list.GetItems(
                    new CamlQuery()
                    {
                        ViewXml = @"<View><Query> 
                                <Where></Where> 
                                </Query></View>"
                    });
                context.Load(items, a => a.IncludeWithDefaultProperties(item => item.File, item => item.File.CheckedOutByUser));
                context.ExecuteQuery();
                foreach (ListItem _item in items)
                {
                    var flag = _item["Description"] != null;
                    if (!flag)
                    {
                        _item["Description"] = "Updated using Timer Job";
                        _item.Update();
                        context.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex) { }
        }
    }
}
