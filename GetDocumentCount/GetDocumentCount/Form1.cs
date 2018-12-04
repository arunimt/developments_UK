using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace GetDocumentCount
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        Utility util = new Utility();
        Constants cons = new Constants();
        SPOUtility spo = new SPOUtility();
        Log_Helper log = new Log_Helper();

        public Form1()
        {
            InitializeComponent();
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            ListItemCollectionPosition position = null;
            int ItemCount = 0;
            int start = 0;
            int end = 0;
            int interval = 100;
            ClientContext context = spo.SPOConnectToken(log, txtUrl.Text);
            try
            {
                if (context != null)
                {
                    var library = context.Web.Lists.GetByTitle(cons.listName);
                    context.Load(library);
                    context.ExecuteQuery();
                    QueryItems(context, library, position, ItemCount);
                    //while(library.ItemCount > end)
                    //{
                    //    var t = GetItemsCountAsync(start, end, interval);
                    //    t.Wait();
                    //    start = start + interval;
                    //    end = end + interval;
                    //}
                }
                lblMsg.Text = "Items with values of " + txtValue.Text + " for " + ddlFilter.Text + " are : " + ItemCount;

            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {

        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            pnlToken.Visible = true;
            pnlCredential.Visible = false;
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            pnlToken.Visible = false;
            pnlCredential.Visible = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pnlToken.Visible = false;
            pnlCredential.Visible = true;
            txtUrl.Text = cons.SiteURL;
            txtUserName.Text = "arun.kumar2@arm.com";
            lblMsg.Text = "";
            LoadFilter();

        }

        private void LoadFilter()
        {
            DataTable dtFilter = util.LoadFilterType();
            for (int i = 0; i < dtFilter.Rows.Count; i++)
            {
                ddlFilter.Items.Add(dtFilter.Rows[i].ItemArray[0]);
            }
        }

        public void QueryItems(ClientContext clientContext, List library, ListItemCollectionPosition position, int ItemCount)
        {

            CamlQuery query = new CamlQuery
            {
                ListItemCollectionPosition = position,
                //ViewXml = @"<View Scope='RecursiveAll'>
                //        <QueryOptions><ViewAttributes Scope='Recursive'/></QueryOptions>
                //        <RowLimit>100</RowLimit>
                //        <Query><Where></Where></Query>
                //    </View>"
                //ViewXml = "<View><Query><Where></Where></Query></View>"
                ViewXml = "<View><RowLimit>1000</RowLimit>" +
                        "<Query><Where><Eq><FieldRef Name='" + ddlFilter.Text + "'/><Value Type='Text'>" + txtValue.Text + "</Value></Eq></Where></Query>" +
                    "</View>"
            };

            var items = library.GetItems(query);
            clientContext.Load(items);
            try
            {
                clientContext.ExecuteQuery();
                ItemCount += items.Count;

                if (items.ListItemCollectionPosition != null)
                {
                    QueryItems(clientContext, library, items.ListItemCollectionPosition, ItemCount);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }
        }

        public async Task GetItemsCountAsync(int start, int end, int interval)
        {
            try
            {
                string sResult = string.Empty;

                start = start + end;
                end = end + interval;
                Uri siteUri = new Uri(txtUrl.Text);
                string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
                string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;

                //const string RESTURL = "{0}/_api/web/lists/GetByTitle('{1}')/items?$skip={2}&$top={3}";
                //string rESTUrl = string.Format(RESTURL, txtUrl.Text, cons.listName, start, end);
                const string RESTURL = "{0}/_api/web/lists/GetByTitle('{1}')/items";
                string rESTUrl = string.Format(RESTURL, txtUrl.Text, cons.listName);

                HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(rESTUrl);
                endpointRequest.Method = "Get";
                endpointRequest.Headers.Add("Authorization", "Bearer " + accessToken);
                endpointRequest.ContentType = "application/json;odata=nometadata";
                endpointRequest.Accept = "application/json;odata=nometadata";
                endpointRequest.Headers.Add("X-RequestDigest", "form digest value");
                endpointRequest.Headers.Add("X-HTTP-Method", "MERGE");
                endpointRequest.Headers.Add("If-Match", "*");

                //WebResponse webResponse = endpointRequest.GetResponse();
                //Stream webStream = webResponse.GetResponseStream();
                //StreamReader responseReader = new StreamReader(webStream);
                //string response = responseReader.ReadToEnd();
                //JObject jobj = JObject.Parse(response);
                //JArray jarr = (JArray)jobj["d"]["results"];

                WebResponse wresp = endpointRequest.GetResponse();
                using (StreamReader sr= new StreamReader(wresp.GetResponseStream()))
                {
                    sResult = sr.ReadToEnd();
                }






                //using (var handler = new HttpClientHandler())
                //{
                //    Uri uri = new Uri(txtUrl.Text);
                //    //Creating HTTP Client
                //    using (var client = new HttpClient(handler))
                //    {
                //        client.DefaultRequestHeaders.Accept.Clear();
                //        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //        //client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/atom+xml"));

                //        //client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
                //        //client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
                //        //client.DefaultRequestHeaders.Add("binaryStringRequestBody", "true");
                //        //client.DefaultRequestHeaders.Add("X-RequestDigest", "form digest value");

                //        HttpResponseMessage response = await client.GetAsync(rESTUrl).ConfigureAwait(false);
                //        //var response = client.GetStringAsync(new Uri(rESTUrl)).Result;
                //        //Ensure 200 (Ok)
                //        response.EnsureSuccessStatusCode();
                //        string resultString = response.Content.ReadAsStringAsync().Result;
                //        var jsonObj = JArray.Parse(resultString);
                //        //ID = jsonObj["ListItemAllFields"]["ID"].ToString();
                //    }
                //}

            }
            catch (Exception ex)
            {
                //return false;
            }
        }
    }
}
