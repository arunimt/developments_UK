using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.SharePoint.Client;

namespace GetLibraryMetadata
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportUtilities util = new ExportUtilities();
            LibraryMetadata library_metadata = new LibraryMetadata();
            int ItemCount = 0;
            string siteUrl = ConfigurationSettings.AppSettings["SiteURL"]; 
            ClientContext clientContext = new ClientContext(siteUrl);
            Web site = clientContext.Web;
            //List sourceList = site.Lists.GetByTitle("ARM Contracts");
            List sourceList = site.Lists.GetByTitle(ConfigurationSettings.AppSettings["LibName"]);
            ListItemCollectionPosition position = null;

            string logMetadata = "Name | Title | DocType | Company | NewCompany | SpecificCompanyName | Requestor | "+
                "EffectiveDate | HardCopyLocation | RevenewEarning | LegalRequest | Comments | SideLetter | CompanyAsText | DocState | " +
                "DocNumber | ApplyHeader | PublishMajorVersion | ContentTypeId | DocNumber | Version | Created | CreatedBy | LastModified | " +
                "ModifiedBy | MainAgreementNumber | DominoAuthor | DominoRequestor | DominoVersion | DocumentSetDescription | EXT | ID";
            util.WriteLine(logMetadata);
            library_metadata.QueryItems(clientContext, sourceList, position, ItemCount);
        }

    }
}
