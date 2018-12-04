<%@ Page Language="C#" Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<WebPartPages:AllowFraming ID="AllowFraming" runat="server" />

<html>
<head>
    <title></title>
    <style>
        #SensitiveDocsTable td, #SensitiveDocsTable th {
            border: 1px solid #ddd;
            padding: 8px;
        }

        #SensitiveDocsTable th {
            background-color: aquamarine;
        }
    </style>
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>

    <%-- Internal blog --%>
    <script type="text/javascript">

        var hostweburl;
        var appweburl;
        var context;
        // Load the required SharePoint libraries
        $(document).ready(function () {
            //Get the URI decoded URLs.
            hostweburl =
                decodeURIComponent(
                    getQueryStringParameter("SPHostUrl")
                );
            appweburl =
                decodeURIComponent(
                    getQueryStringParameter("SPAppWebUrl")
                );

            listsToQuery =
                decodeURIComponent(
                    getQueryStringParameter("SensitiveViewLists")
                );
            // resources are in URLs in the form:
            // web_url/_layouts/15/resource
            var scriptbase = hostweburl + "/_layouts/15/";
            // Load the js files and continue to the successHandler
            $.getScript(scriptbase + "SP.RequestExecutor.js", execOperation);
          //  $.getScript(scriptbase + "SP.RequestExecutor.js", operation);

          //  $.getScript(scriptbase + "init.js",
          //  function () {
          //      $.getScript(scriptbase + "SP.Runtime.js",
          //         function () {
          //      $.getScript(scriptbase + "SP.js", operation);
          //    });
          //});

        });

        function operation() {
            //alert("Hi");
            //var ctx = new SP.ClientContext(appweburl),
            //    factory = new SP.ProxyWebRequestExecutorFactory(appweburl),
            //    web;
            var ctx = new SP.ClientContext.get_current();
            //ctx.set_webRequestExecutorFactory(factory);
            web = ctx.get_web();
            ctx.load(web);
            ctx.executeQueryAsync(function () {
                // log the name of the app web to the console
                console.log(web.get_title());
            }, function (sender, args) {
                console.log("Error : " + args.get_message());
            });


        }

        // Function to prepare and issue the request to get
        //  SharePoint data
        //function execCrossDomainRequest() {
        function execOperation() {
            //context = new SP.ClientContext(appweburl);
            //alert(context);
            // executor: The RequestExecutor object
            // Initialize the RequestExecutor with the app web URL.
            var executor = new SP.RequestExecutor(appweburl);
            var lists = [];
            if (listsToQuery != "") {
                lists = listsToQuery.split(',');
            }
            //var lists = ["Draft", "Issued", "Active", "Expired"];
            debugger;
            var i;
            for (i = 0; i < lists.length; i++) {
                //var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('" + lists[i] + "')/Items?$select=Title,ContractID,EffectiveDate,ExpiryDate,DocumentType,Counterparty,MasterAgreementNumber,Perpetual,Requester,ServiceNowRequestNumber,FileLeafRef,FileSystemObjectType&$filter=Sensitive eq 'yes'";
                //var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('" + lists[i] + "')/Items?$select=Title,ContractID,EffectiveDate,ExpiryDate,DocumentType,Counterparty,MasterAgreementNumber,Perpetual,Requester,ServiceNowRequestNumber,FileLeafRef,FileSystemObjectType&$filter=Sensitive eq 'yes' and Title eq 'CDS (Active)'";
                //var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('" + lists[i] + "')/Items?$select=Title,ContractID,EffectiveDate,ExpiryDate,DocumentType,Counterparty,MasterAgreementNumber,Perpetual,Requester,ServiceNowRequestNumber,FileLeafRef,FileSystemObjectType";
                var restUrl = appweburl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('" + lists[i] + "')/items?@target='" + hostweburl + "/" + lists[i] + "'&$select=Title,ContractID,EffectiveDate,ExpiryDate,DocumentType,Counterparty,MasterAgreementNumber,Perpetual,Requester,ServiceNowRequestNumber,FileLeafRef,FileSystemObjectType,LibraryType&$filter=Sensitive eq 'yes'";
                console.log(restUrl);

                //var oList = clientContext.get_web().get_lists().getByTitle(lists[i]);

                 //alert(oList);

                // Issue the call against the app web.
                // To get the title using REST we can hit the endpoint:
                //      appweburl/_api/web/lists/getbytitle('listname')/items
                // The response formats the data in the JSON format.
                // The functions successHandler and errorHandler attend the
                //      sucess and error events respectively.
                executor.executeAsync(
                    {
                        url: restUrl,
                        method: "GET",
                        headers: { "Accept": "application/json; odata=verbose" },
                        success: successHandler,
                        error: errorHandler
                    }
                );
            }
        }

        // Function to handle the success event.
        // Prints the data to the page.
        function successHandler(data) {
            var jsonObject = JSON.parse(data.body);
            var blogsHTML = "";
            var results = jsonObject.d.results;
            for (var i = 0; i < results.length; i++) {
                if (results[i].FileSystemObjectType == 0) {
                    var metadata = [];
                    metadata.push(results[i].FileLeafRef);
                    metadata.push(results[i].Title);
                    metadata.push(results[i].ContractID);
                    metadata.push(results[i].DocumentType);
                    metadata.push(results[i].Counterparty);
                    metadata.push(results[i].Perpetual);
                    if (results[i].EffectiveDate != null) {
                        metadata.push(GetFormattedDate(results[i].EffectiveDate));
                    }
                    else {
                        metadata.push("");
                    }
                    if (results[i].ExpiryDate != null) {
                        metadata.push(GetFormattedDate(results[i].ExpiryDate));
                    }
                    else {
                        metadata.push("");
                    }
                    metadata.push(results[i].MasterAgreementNumber);
                    metadata.push(results[i].Requester);
                    metadata.push(results[i].ServiceNowRequestNumber);
                    metadata.push(results[i].LibraryType);
                    console.log(metadata);
                    var html = '<tr><td>' + metadata[0] + '</td><td>' + metadata[1] + '</td><td>' + metadata[2] + '</td><td>';
                    var html = html + metadata[3] + '</td><td>' + metadata[4] + '</td><td>' + metadata[5] + '</td><td>' + metadata[6] + '</td><td>';
                    var html = html + metadata[7] + '</td><td>' + metadata[8] + '</td><td>' + metadata[9] + '</td><td>' + metadata[10] + '</td><td>' + metadata[11] + '</td></tr>';
                    $("#SensitiveDocsTable").find('tbody').append(html);
                }
            }

            //$('#internal').append(blogsHTML);
        }

        // Function to handle the error event.
        // Prints the error message to the page.
        function errorHandler(data, errorCode, errorMessage) {
            //document.getElementById("internal").innerText =
              //  "Could not complete cross-domain call: " + errorMessage;
        }

        // Function to retrieve a query string value.
        // For production purposes you may want to use
        //  a library to handle the query string.
        function getQueryStringParameter(paramToRetrieve) {
            var params =
                document.URL.split("?")[1].split("&");
            var strParams = "";
            for (var i = 0; i < params.length; i = i + 1) {
                var singleParam = params[i].split("=");
                if (singleParam[0] == paramToRetrieve)
                    return singleParam[1];
            }
        }

        function GetFormattedDate(date) {
            var todayTime = new Date(date);
            var month = todayTime.getMonth() + 1;
            if (month < 10) {
                month = '0' + month;
            }
            var day = todayTime.getDate();
            if (day < 10) {
                day = '0' + day;
            }
            var year = todayTime.getFullYear();
            return year + "-" + month + "-" + day;
        }
    </script>

</head>
<body>
    <div>
        <table id="SensitiveDocsTable">
            <thead>
                <tr>
                    <td>FileName</td>
                    <td>Title</td>
                    <td>ContractID</td>
                    <td>DocumentType</td>
                    <td>Counterparty</td>
                    <td>Perpetual</td>
                    <td>Effective Date</td>
                    <td>Expiry Date</td>
                    <td>MasterAgreementNumber</td>
                    <td>Requester</td>
                    <td>ServiceNow Request Number</td>
                    <td>Library</td>
                </tr>
            </thead>
            <tbody>
            </tbody>
        </table>
    </div>
</body>
</html>
