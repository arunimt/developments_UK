using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace GetLibraryMetadata
{
    class LibraryMetadata
    {
        #region Properties
        string _Name;
        string _Title;
        string _DocType;
        string _Company;
        string _NewCompany;
        string _SpecificCompanyName;
        string _Requestor;
        string _DocState;
        Nullable<DateTime> _EffectiveDate;
        string _HardCopyLocation;
        string _RevenewEarning;
        string _LegalRequest;
        string _Comments;
        Nullable<int> _SideLetter;
        string _CompanyAsText;
        string _DocNumber;
        Nullable<bool> _ApplyHeader;
        Nullable<bool> _PublishMajorVersion;
        string _ContentTypeId;
        string _Version;
        Nullable<DateTime> _Created;
        Nullable<DateTime> _lastModified;
        string _CreatedBy;
        string _ModifiedBy;
        string _MainAgreementNumber;
        string _DominoAuthor;
        string _DominoRequestor;
        string _DominoVersion;
        string _DocumentSetDescription;
        string _EXT;
        string _ID;

        public string Name { get => _Name; set => _Name = value; }
        public string Title { get => _Title; set => _Title = value; }
        public string DocType { get => _DocType; set => _DocType = value; }
        public string Company { get => _Company; set => _Company = value; }
        public string NewCompany { get => _NewCompany; set => _NewCompany = value; }
        public string SpecificCompanyName { get => _SpecificCompanyName; set => _SpecificCompanyName = value; }
        public string Requestor { get => _Requestor; set => _Requestor = value; }
        public string DocState { get => _DocState; set => _DocState = value; }
        public DateTime? EffectiveDate { get => _EffectiveDate; set => _EffectiveDate = value; }
        public string HardCopyLocation { get => _HardCopyLocation; set => _HardCopyLocation = value; }
        public string LegalRequest { get => _LegalRequest; set => _LegalRequest = value; }
        public string RevenewEarning { get => _RevenewEarning; set => _RevenewEarning = value; }
        public string Comments { get => _Comments; set => _Comments = value; }
        public int? SideLetter { get => _SideLetter; set => _SideLetter = value; }
        public string CompanyAsText { get => _CompanyAsText; set => _CompanyAsText = value; }
        public string DocNumber { get => _DocNumber; set => _DocNumber = value; }
        public bool? ApplyHeader { get => _ApplyHeader; set => _ApplyHeader = value; }
        public bool? PublishMajorVersion { get => _PublishMajorVersion; set => _PublishMajorVersion = value; }
        public string ContentTypeId { get => _ContentTypeId; set => _ContentTypeId = value; }
        public string Version { get => _Version; set => _Version = value; }
        public DateTime? Created { get => _Created; set => _Created = value; }
        public DateTime? LastModified { get => _lastModified; set => _lastModified = value; }
        public string CreatedBy { get => _CreatedBy; set => _CreatedBy = value; }
        public string ModifiedBy { get => _ModifiedBy; set => _ModifiedBy = value; }
        public string MainAgreementNumber { get => _MainAgreementNumber; set => _MainAgreementNumber = value; }
        public string DominoAuthor { get => _DominoAuthor; set => _DominoAuthor = value; }
        public string DominoRequestor { get => _DominoRequestor; set => _DominoRequestor = value; }
        public string DominoVersion { get => _DominoVersion; set => _DominoVersion = value; }
        public string DocumentSetDescription { get => _DocumentSetDescription; set => _DocumentSetDescription = value; }
        public string EXT { get => _EXT; set => _EXT = value; }
        public string ID { get => _ID; set => _ID = value; }

        #endregion

        public void QueryItems(ClientContext clientContext, List library, ListItemCollectionPosition position, int ItemCount)
        {
            ExportUtilities util = new ExportUtilities();
            LogUtility log = new LogUtility();
            CamlQuery query = new CamlQuery
            {
                ListItemCollectionPosition = position,
                ViewXml = @"<View Scope='RecursiveAll'>
                        <QueryOptions><ViewAttributes Scope='Recursive'/></QueryOptions>
                        <RowLimit>100</RowLimit>
                        <Query><Where></Where></Query>
                    </View>"
            };

            var items = library.GetItems(query);
            clientContext.Load(items);
            try
            {
                clientContext.ExecuteQuery();
                //var start = items.First();
                //var end = items.First();
                ItemCount += items.Count;
                log.WriteLine("Total Items Processed: " + ItemCount);
                log.WriteLine("Total count in this Iteration: " + items.Count);
                log.WriteLine("Start ID: " + items[0]["ID"].ToString());
                log.WriteLine("End ID: " + items[99]["ID"].ToString());

                foreach (var item in items)
                {
                    LibraryMetadata metadataInfo = new LibraryMetadata();

                    try
                    {
                        metadataInfo.Name = item.FieldValues["FileLeafRef"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.Name = null;
                    }
                    try
                    {
                        metadataInfo.Title = item.FieldValues["Title"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.Title = null;
                    }
                    try
                    {
                        metadataInfo.DocType = item.FieldValues["DocType"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DocType = null;
                    }
                    try
                    {
                        metadataInfo.Company = item.FieldValues["CompanyAsText"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.Company = null;
                    }
                    try
                    {
                        metadataInfo.NewCompany = item.FieldValues["New_x0020_Company"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.NewCompany = null;
                    }
                    try
                    {
                        metadataInfo.SpecificCompanyName = item.FieldValues["Specific_x0020_Company_x0020_Name"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.SpecificCompanyName = null;
                    }
                    try
                    {
                        metadataInfo.Requestor = item.FieldValues["Requestor"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.Requestor = null;
                    }
                    try
                    {
                        metadataInfo.EffectiveDate = DateTime.Parse(item.FieldValues["Effective_x0020_Date"].ToString());
                    }
                    catch (Exception)
                    {
                        metadataInfo.EffectiveDate = null;
                    }
                    try
                    {
                        metadataInfo.HardCopyLocation = item.FieldValues["Hardcopy_x0020_Location"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.HardCopyLocation = null;
                    }
                    try
                    {
                        metadataInfo.RevenewEarning = item.FieldValues["Revenue_x0020_Earning"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.RevenewEarning = null;
                    }
                    try
                    {
                        metadataInfo.LegalRequest = item.FieldValues["Legal_x0020_Request"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.LegalRequest = null;
                    }
                    try
                    {
                        metadataInfo.Comments = item.FieldValues["Comments"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.Comments = null;
                    }
                    try
                    {
                        metadataInfo.SideLetter = int.Parse(item.FieldValues["Side_x0020_Letters"].ToString());
                    }
                    catch (Exception)
                    {
                        metadataInfo.SideLetter = null;
                    }
                    try
                    {
                        metadataInfo.CompanyAsText = item.FieldValues["CompanyAsText"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.CompanyAsText = null;
                    }
                    try
                    {
                        metadataInfo.DocState = item.FieldValues["DocState"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DocState = null;
                    }
                    try
                    {
                        metadataInfo.DocNumber = item.FieldValues["DocNumber"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DocNumber = null;
                    }
                    try
                    {
                        metadataInfo.ApplyHeader = bool.Parse(item.FieldValues["ApplyHeader"].ToString());
                    }
                    catch (Exception)
                    {
                        metadataInfo.ApplyHeader = null;
                    }
                    try
                    {
                        metadataInfo.PublishMajorVersion = bool.Parse(item.FieldValues["PublishMajorVersion"].ToString());
                    }
                    catch (Exception)
                    {
                        metadataInfo.PublishMajorVersion = null;
                    }
                    try
                    {
                        metadataInfo.ContentTypeId = item.FieldValues["ContentTypeId"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.ContentTypeId = null;
                    }
                    try
                    {
                        metadataInfo.Version = item.FieldValues["_UIVersionString"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.Version = null;
                    }
                    try
                    {
                        metadataInfo.Created = DateTime.Parse(item.FieldValues["Created"].ToString());
                    }
                    catch (Exception)
                    {
                        metadataInfo.Created = null;
                    }
                    try
                    {
                        metadataInfo.CreatedBy = item.FieldValues["Created_x0020_By"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.CreatedBy = null;
                    }
                    try
                    {
                        metadataInfo.LastModified = DateTime.Parse(item.FieldValues["Modified"].ToString());
                    }
                    catch (Exception)
                    {
                        metadataInfo.LastModified = null;
                    }
                    try
                    {
                        metadataInfo.ModifiedBy = item.FieldValues["Modified_x0020_By"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.ModifiedBy = null;
                    }
                    try
                    {
                        metadataInfo.MainAgreementNumber = item.FieldValues["Original_x0020_Agreement_x0020_Title"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.MainAgreementNumber = null;
                    }
                    try
                    {
                        metadataInfo.DominoAuthor = item.FieldValues["Domino_x0020_Author"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DominoAuthor = null;
                    }
                    try
                    {
                        metadataInfo.DominoRequestor = item.FieldValues["Domino_x0020_Requestor"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DominoRequestor = null;
                    }
                    try
                    {
                        metadataInfo.DominoVersion = item.FieldValues["Current_x0020_Version"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DominoVersion = null;
                    }
                    try
                    {
                        metadataInfo.DocumentSetDescription = item.FieldValues["DocumentSetDescription"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.DocumentSetDescription = null;
                    }
                    try
                    {
                        metadataInfo.EXT = item.FieldValues["EXT"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.EXT = null;
                    }
                    try
                    {
                        metadataInfo.ID = item.FieldValues["ID"].ToString();
                    }
                    catch (Exception)
                    {
                        metadataInfo.ID = null;
                    }

                    //Name = CheckNull("FileLeafRef", item),
                    //Title = CheckNull("Title", item),
                    //DocType = CheckNull("DocType", item),
                    //Company = CheckNull("CompanyAsText", item),
                    //NewCompany = CheckNull("New_x0020_Company", item),
                    //SpecificCompanyName = CheckNull("Specific_x0020_Company_x0020_Name", item),
                    //Requestor = CheckNull("Requestor", item),
                    //EffectiveDate = CheckNullDate("Effective_x0020_Date", item),
                    //HardCopyLocation = CheckNull("Hardcopy_x0020_Location", item),
                    //RevenewEarning = CheckNull("Revenue_x0020_Earning", item),
                    //LegalRequest = CheckNull("Legal_x0020_Request", item),
                    //Comments = CheckNull("Comments", item),
                    //SideLetter = CheckNullInt("Side_x0020_Letters", item),
                    //CompanyAsText = CheckNull("CompanyAsText", item),
                    //DocState = CheckNull("DocState", item),
                    //DocNumber = CheckNull("DocNumber", item),
                    //ApplyHeader = CheckNullBool("ApplyHeader", item),
                    //PublishMajorVersion = CheckNullBool("PublishMajorVersion", item),
                    //ContentTypeId = CheckNull("ContentTypeId", item),
                    //Version = CheckNull("_UIVersionString", item),
                    //Created = CheckNullDate("Created", item),
                    //CreatedBy = CheckNull("Created_x0020_By", item),
                    //LastModified = CheckNullDate("Modified", item),
                    //ModifiedBy = CheckNull("Modified_x0020_By", item),
                    //MainAgreementNumber = CheckNull("Original_x0020_Agreement_x0020_Title", item),
                    //DominoAuthor = CheckNull("Domino_x0020_Author", item),
                    //DominoRequestor = CheckNull("Domino_x0020_Requestor", item),
                    //DominoVersion = CheckNull("Current_x0020_Version", item),
                    //DocumentSetDescription = CheckNull("DocumentSetDescription", item),
                    //EXT = CheckNull("EXT", item),
                    //ID = CheckNull("ID", item)
                    //};


                    if (metadataInfo.ContentTypeId.ToString().StartsWith("0x0120D520"))
                    {
                        string logMetadata = metadataInfo.Name + " | " + metadataInfo.Title + " | " + metadataInfo.DocType + " | " + metadataInfo.Company + " | " + metadataInfo.NewCompany + " | " +
                            metadataInfo.SpecificCompanyName + " | " + metadataInfo.Requestor + " | " + metadataInfo.EffectiveDate + " | " + metadataInfo.HardCopyLocation + " | " + metadataInfo.RevenewEarning + " | " +
                            metadataInfo.LegalRequest + " | " + metadataInfo.Comments + " | " + metadataInfo.SideLetter + " | " + metadataInfo.CompanyAsText + " | " + metadataInfo.DocState + " | " + metadataInfo.DocNumber + " | " +
                            metadataInfo.ApplyHeader + " | " + metadataInfo.PublishMajorVersion + " | " + metadataInfo.ContentTypeId + " | " + metadataInfo.DocNumber + " | " + metadataInfo.Version + " | " + metadataInfo.Created + " | " +
                            metadataInfo.CreatedBy + " | " + metadataInfo.LastModified + " | " + metadataInfo.ModifiedBy + " | " + metadataInfo.MainAgreementNumber + " | " + metadataInfo.DominoAuthor + " | " +
                            metadataInfo.DominoRequestor + " | " + metadataInfo.DominoVersion + " | " + metadataInfo.DocumentSetDescription + " | " + metadataInfo.EXT + " | " + metadataInfo.ID;


                        util.WriteLine(logMetadata);
                    }

            }
                if (items.ListItemCollectionPosition != null)
            {
                QueryItems(clientContext, library, items.ListItemCollectionPosition, ItemCount);
            }
        }
            catch (Exception ex)
            {
                log.WriteLine(ex.Message);
            }
}

public string CheckNull(string internal_name, ListItem item)
{
    string value = "";
    if (item.FieldValues[internal_name] == null)
        value = null;
    else
        value = item.FieldValues[internal_name].ToString();

    return value;
}
public Nullable<DateTime> CheckNullDate(string internal_name, ListItem item)
{
    Nullable<DateTime> value;
    if (item.FieldValues[internal_name] == null)
        value = null;
    else
        value = DateTime.Parse(item.FieldValues[internal_name].ToString());

    return value;
}

public Nullable<int> CheckNullInt(string internal_name, ListItem item)
{
    Nullable<int> value;
    if (item.FieldValues[internal_name] == null)
        value = null;
    else
        value = int.Parse(item.FieldValues[internal_name].ToString());

    return value;
}

public Nullable<bool> CheckNullBool(string internal_name, ListItem item)
{
    Nullable<bool> value;
    if (item.FieldValues[internal_name] == null)
        value = null;
    else
        value = bool.Parse(item.FieldValues[internal_name].ToString());

    return value;
}
        
    }
}
