using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertDocSetToFolder
{
    class Program
    {
        static void Main(string[] args)
        {
            Utility util = new Utility();
            //ClientContext context = util.SPOConnectToken();
            ClientContext context = util.SPOConnect();
            ListItemCollectionPosition position = null;

            if (context != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                var list = context.Web.Lists.GetByTitle(ConfigurationSettings.AppSettings["DocLibName"]);
#pragma warning restore CS0618 // Type or member is obsolete
                context.Load(list);
                context.Load(list.RootFolder.Folders);
                context.ExecuteQuery();

                DocSetToFolder(context, list, position);
            }
        }

        static public ContentTypeId GetFolderCT(ClientContext context)
        {
            ContentTypeId folderCT = new ContentTypeId();

            var ctTypes = context.Web.AvailableContentTypes;
            context.Load(ctTypes);
            context.ExecuteQuery();

            foreach (var ctType in ctTypes)
            {
                if(ctType.Name == "Folder")
                {
                    folderCT = ctType.Id;
                    break;
                }
            }

            return folderCT;
        }

        static public void DocSetToFolder(ClientContext clientContext, List library, ListItemCollectionPosition position)
        {
            CamlQuery query = new CamlQuery
            {
                ListItemCollectionPosition = position,
               ViewXml = @"<View Scope='RecursiveAll'><RowLimit>" + 1000 + "</RowLimit><Query></Query></View>"
            };

#pragma warning disable CS0618 // Type or member is obsolete
            var folder = GetFolderCT(clientContext);
#pragma warning restore CS0618 // Type or member is obsolete
            var items = library.GetItems(query);
            clientContext.Load(items);
            try
            {
                clientContext.ExecuteQuery();
                foreach (var item in items)
                {
                    if (item["HTML_x0020_File_x0020_Type"] != null)
                    {
                        item["ContentTypeId"] = folder;
                        item["xd_ProgID"] = null;
                        item["HTML_x0020_File_x0020_Type"] = null;
                        item.Update();
                        clientContext.ExecuteQuery();
                    }
                    
                }
                if (items.ListItemCollectionPosition != null)
                {
                    DocSetToFolder(clientContext, library, items.ListItemCollectionPosition);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
