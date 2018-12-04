using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SensitiveViewAddInWeb.Pages
{
    public partial class ViewSensitive : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            divError.Visible = false;
            string filter = ddlFilter.SelectedValue;
            string strSearch = txtSearch.Text;
            LoadData(filter, strSearch);
        }

        protected void BtnSearch_Click(object sender, EventArgs e)
        {
            string filter = ddlFilter.SelectedValue;
            string strSearch = txtSearch.Text;
            LoadData(filter,strSearch);
        }
      
        protected void LoadData(string filter, string strSearch)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
            DataTable dt = new DataTable();
            dt.Columns.Add("FileName");
            dt.Columns.Add("Title");
            dt.Columns.Add("ContractID");
            dt.Columns.Add("DocumentType");
            dt.Columns.Add("Counterparty");
            dt.Columns.Add("Perpetual");
            dt.Columns.Add("Effective Date");
            dt.Columns.Add("Expiry Date");
            dt.Columns.Add("MasterAgreementNumber");
            dt.Columns.Add("Requester");
            dt.Columns.Add("ServiceNow Request Number");
            dt.Columns.Add("Library");

            using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
            {
                if (Context.Request.QueryString["SensitiveViewLists"] != "")
                {
                    divSensitive.Visible = true;
                    var lists = Context.Request.QueryString["SensitiveViewLists"];
                    var listArr = lists.Split(',');

                    foreach (var list in listArr)
                    {
                        clientContext.Load(clientContext.Web, a => a.Lists);
                        clientContext.ExecuteQuery();

                        List _list = clientContext.Web.Lists.GetByTitle(list);
                        var items = _list.GetItems(new CamlQuery() { ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><IsNotNull><FieldRef Name=\"File_x0020_Type\" /></IsNotNull></Where></Query></View>" });
                        clientContext.Load(items);
                        clientContext.ExecuteQuery();
                        
                        foreach (var item in items)
                        {
                            switch (filter)
                            {
                                case "FileName":
                                    if (item["FileRef"] != null)
                                    {
                                        if (item["FileRef"].ToString().Split('/').LastOrDefault().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item,dt);                
                                        }
                                    }
                                    break;
                                case "Title":
                                    if (item["Title"] != null)
                                    {
                                        if (item["Title"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "ContractID":
                                    if (item["ContractID"] != null)
                                    {
                                        if (item["ContractID"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "DocumentType":
                                    if (item["DocumentType"] != null)
                                    {
                                        if (item["DocumentType"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "Counterparty":
                                    if (item["Counterparty"] != null)
                                    {
                                        if (item["Counterparty"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "Perpetual":
                                    if (item["Perpetual"] != null)
                                    {
                                        if (item["Perpetual"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "Effective Date":
                                    if (item["EffectiveDate"] != null)
                                    {
                                        if (item["EffectiveDate"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "Expiry Date":
                                    if (item["ExpiryDate"] != null)
                                    {
                                        if (item["ExpiryDate"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "MasterAgreementNumber":
                                    if (item["MasterAgreementNumber"] != null)
                                    {
                                        if (item["MasterAgreementNumber"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "Requester":
                                    if (item["Requester"] != null)
                                    {
                                        if (item["Requester"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "ServiceNow Request Number":
                                    if (item["ServiceNowRequestNumber"] != null)
                                    {
                                        if (item["ServiceNowRequestNumber"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                case "Library":
                                    if (item["LibraryType"] != null)
                                    {
                                        if (item["LibraryType"].ToString().Contains(txtSearch.Text))
                                        {
                                            BuildDataTable(item, dt);
                                        }
                                    }
                                    break;
                                default:
                                    BuildDataTable(item, dt);
                                    break;
                            }
                            
                        }
                    }

                    GridSensitiveView.DataSource = dt;
                    GridSensitiveView.DataBind();
                }
                else
                {
                    divSensitive.Visible = false;
                    divError.Visible = true;
                    lblError.Text = "Enter the SensitiveViewLists property by editing the webpart";
                }
            }
        }

        protected void BuildDataTable(Microsoft.SharePoint.Client.ListItem item, DataTable dt)
        {
            string[] doc = new string[12];
            doc[0] = item["FileRef"]?.ToString().Split('/').LastOrDefault();
            doc[1] = item["Title"]?.ToString();
            doc[2] = item["ContractID"]?.ToString();
            doc[3] = item["DocumentType"]?.ToString();
            doc[4] = item["Counterparty"]?.ToString();
            doc[5] = item["Perpetual"]?.ToString();
            if (item["EffectiveDate"] != null)
            {
                string _date = item["EffectiveDate"].ToString();
                DateTime _dt = DateTime.Parse(_date);
                doc[6] = _dt.ToShortDateString();
            }
            else
            {
                doc[6] = null;
            }
            if (item["ExpiryDate"] != null)
            {
                string _date = item["ExpiryDate"].ToString();
                DateTime _dt = DateTime.Parse(_date);
                doc[7] = _dt.ToShortDateString();
            }
            else
            {
                doc[7] = null;
            }
            doc[8] = item["MasterAgreementNumber"]?.ToString();
            doc[9] = item["Requester"]?.ToString();
            doc[10] = item["ServiceNowRequestNumber"]?.ToString();
            doc[11] = item["LibraryType"]?.ToString();

            dt.Rows.Add(doc);
        }

        protected void GridSensitiveView_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridSensitiveView.PageIndex = e.NewPageIndex;
            string filter = ddlFilter.SelectedValue;
            string strSearch = txtSearch.Text;
            LoadData(filter, strSearch);
        }
    }
}