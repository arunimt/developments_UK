using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateLibraryType
{
    class Program
    {

        static void Main(string[] args)
        {
            Utility util = new Utility();
            ClientContext context = util.SPOConnect();

            try
            {
                if (context != null)
                {
                    var list = context.Web.Lists.GetByTitle(ConfigurationSettings.AppSettings["DocLibName"]);
                    context.Load(list);
                    context.Load(list.RootFolder.Folders);
                    context.ExecuteQuery();
                    //cntrlr.Processfiles(context, ConfigurationSettings.AppSettings["SourceLocation"], log);

                    var items = list.GetItems(
                        new CamlQuery()
                        {
                            ViewXml = @"<View Scope='RecursiveAll'><Query> 
                                <Where><IsNotNull><FieldRef Name='File_x0020_Type' /></IsNotNull></Where> 
                                </Query></View>"
                        });
                    context.Load(items, a => a.IncludeWithDefaultProperties(item => item.File, item => item.File.CheckedOutByUser));
                    context.ExecuteQuery();
                    foreach (var item in items)
                    {
                        if (item.File.CheckOutType != CheckOutType.None)
                        {
                            item.File.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                            item.File.CheckOut();
                            ListItem _item = item.File.ListItemAllFields;
                            _item["LibraryType"] = ConfigurationSettings.AppSettings["DocLibName"];
                            _item.Update();
                            item.File.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                        }
                    }

                    foreach (Folder subFolder in list.RootFolder.Folders)
                    {
                        try
                        {
                            context.Load(subFolder.Files);
                            context.ExecuteQuery();
                            foreach (File f in subFolder.Files)
                            {
                                try
                                {
                                    if (f.CheckOutType == CheckOutType.None)
                                    {
                                        f.CheckOut();
                                        ListItem item = f.ListItemAllFields;
                                        item["LibraryType"] = ConfigurationSettings.AppSettings["DocLibName"];
                                        item.Update();
                                        f.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                                    }
                                    else
                                    {
                                        f.CheckOut();
                                        ListItem item = f.ListItemAllFields;
                                        item["LibraryType"] = ConfigurationSettings.AppSettings["DocLibName"];
                                        item.Update();
                                        f.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.Write(ex.Message);
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //log.Error(ex.Message);
            }
        }
    }
}
