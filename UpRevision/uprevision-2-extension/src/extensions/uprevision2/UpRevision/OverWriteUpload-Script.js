/*
Description: The file performs an overwrite action on an existing document.
The following actions are performed:
	a) Checks if the file exists
	b) Checks if the file is checked out
	c) If not checked-Out, it triggers the check-out
	d) Uploads the document on SharePoint
	c) Checks-in the document with comment
*/

//Global Variables
var contractId = "";
var fileName = "";
var folderName = "";
var serverRelativUrl = "";
var metadataRelativeUrl = "";
var UpdateItemId = "";
var contextURL = _spPageContextInfo.webAbsoluteUrl;
var deferred = $.Deferred();
var version = "";
var dataArr;
var sameUser = false;

//process flow begins here
var OverWriteUpload = function () {
    $('#metadata').hide();
    $('#detail').hide();
    $('#message').hide();
    if($('#ContactId').val() == ""){
        var msgText = "Contract ID cannot be blank. Please enter a Contract ID and try again.";
        $('#msg').html(msgText).css('color', '#e81123');
        $('#message').show();
    }
    else{
        var msgText = "Contract Search is in progress.";
        $('#progressMsg').html(msgText).css('color', '#107c10');
        $('#progress').show();
        checkIfFileExistsInDraft();
    }
}

//confirm metadata and check for "is current user admin user"
var ConfirmYes = function () {
    getCurrentUserPermission();
    $('#detail').hide();
    $('#buttonmenu').show();
    $('#uploadfile').show();
}

//Get current Logged in user for already checkout file
function getCurrentUser() {
    // alert("get Current user");
    $.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/CurrentUser",
        method: "GET",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function (data) {
            // alert(data.d.Id);
            console.log('current user: ' + data.d.Title); 
            checkIfThefileIsCheckedOut(data.d.Id);
            deferred.resolve();
        },
        error: function () {
            // alert("error");
            console.log(error.responseJSON.error.message.value);
            deferred.resolve();
        }
    });
    return deferred.promise();
}

//Get current logged in user member of 
function getCurrentUserPermission() {
    //alert("getCurrentUserMetadata");
    $.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/sitegroups/getbyname('Commercial-Admin')/CanCurrentUserViewMembership",
        method: "GET",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function (data) {
            if(data.d.CanCurrentUserViewMembership){
                console.log('current user is member of Commercial-Admin');
                $('#editMetadata').css('display', 'inline');
                $('#editMetadata').show(); 
            }
            else{
                console.log('current user is not a member of Commercial-Admin');
                $('#editMetadata').hide();
            }
            deferred.resolve();
        },
        error: function (error) {
            alert(error.responseJSON.error.message.value);
            deferred.resolve();
        }
    });
    return deferred.promise();
}

//validation for only .docs files
var CheckFileType = function (){
    fileName = $('#file').val();
    fileName = fileName.substring(fileName.lastIndexOf('\\') + 1);
    // alert(fileName);
    fileExt = fileName.split('.')[1];
    // alert(fileExt);
    folderName = fileName.substring(0, fileName.indexOf('.'));
    // alert(folderName);
    if(fileExt == "docx")
    {
        if(folderName.toUpperCase() == contractId.toUpperCase()){
            $('#message').hide();
            var confirmMsgText = "'Do you wish to upload this document as the next version of " + contractId + " in SharePoint ?";
            $('#confirmUploadMsg').html(confirmMsgText);
            $('#confirmupload').show();
        }
        else{
            var msgText = "File name should be same as contract Id. Please rename the file and then try again.";
            $('#msg').html(msgText).css('color', '#e81123');
            $('#message').show();
        }
    }
    else
    {
        var msgText = "Only .docx are valid file types. Please select a different file and try again.";
        $('#msg').html(msgText).css('color', '#e81123');
        $('#message').show();
    }
}

//Get formatted Effective Date
function GetFormattedDate(effectivedate) {
    var todayTime = new Date(effectivedate);
    var month = todayTime.getMonth()+1;
    if (month < 10){
        month = '0' + month;
    }
    var day = todayTime.getDate();
    if (day < 10)
    {
        day = '0' + day;
    }
    var year = todayTime.getFullYear();
    return year + "-" + month + "-" + day;
}

//upload document to SharePoint
var ConfirmUpload = function () {
    if($('#comment').val().length == 0){
        var msgText = "Check-in comments need to be present.";
        $('#lblErrMsg').html(msgText).css('color', '#e81123');
    }
    else if($('#comment').val().length > 255){
        var msgText = "Comment cannot be more that 255 characters.";
        $('#lblErrMsg').html(msgText).css('color', '#e81123');
    }
    else{
        var msgText = "";
        $('#lblErrMsg').html(msgText).css('color', '#e81123');
        $('#buttonmenu').hide();
        $('#confirmupload').hide();
        if(sameUser){
            uploadFileSync();
        }
        else{
            CheckOut();
        }
    }
}

//click of Return to main menu
var ReturnToMainMenu = function () {
    $('#ContactId').val("");
    $('#search').show();
    $('#metadata').hide();
    $('#detail').hide();
    $('#buttonmenu').hide();
    $('#uploadfile').hide();
    $('#confirmupload').hide();
    $('#message').hide();
    $('#file').val("");
    $('#comment').val("");
    $('#success').hide();
}

// update metadata
var SaveMetaData = function () {
    var valid = true;
    if($('#CounterPartyEdit').val().length > 255 ){
        valid = false;
        var msgText = "Counter Party cannot be more that 255 characters.";
        $('#lblCounterPartyErrMsg').html(msgText).css('color', '#e81123');        
    }
    if($('#DocumentTypeEdit').val().length > 255){
        valid = false;
        var msgText = "Document Type cannot be more that 255 characters.";
        $('#lblDocumentTypeErrMsg').html(msgText).css('color', '#e81123');
    }
    if($('#DocumentTitleEdit').val().length > 255 ){
        valid = false;
        var msgText = "Document Title cannot be more that 255 characters.";
        $('#lblDocumentTitleErrMsg').html(msgText).css('color', '#e81123');
    }
    if($('#MasterAgreementNumEdit').val().length > 255 ){
        valid = false;
        var msgText = "Master Agreement Number cannot be more that 255 characters.";
        $('#lblMasterAgreementNumErrMsg').html(msgText).css('color', '#e81123');
    }    
    if(valid){
        if(sameUser){
            SaveData();
            DisableMetadata();
        }
        else{
            MetaDataCheckOut();
            DisableMetadata();
        }
    }
}

var DisableMetadata = function () {
    $("#CounterPartyEdit").attr('disabled', 'true');             
    $("#DocumentTypeEdit").attr('disabled', 'true');           
    $("#DocumentTitleEdit").attr('disabled', 'true');             
    $("#EffectiveDateEdit ").attr('disabled', 'true');  
    $("#MasterAgreementNumEdit").attr('disabled', 'true');  

    $("#CounterPartyEdit").css("background-color", "#EBEBE4");
    $("#DocumentTypeEdit").css("background-color", "#EBEBE4");    
    $("#DocumentTitleEdit").css("background-color", "#EBEBE4");
    $("#EffectiveDateEdit").css("background-color", "#EBEBE4");
    $("#MasterAgreementNumEdit").css("background-color", "#EBEBE4");

    msgText = "";
    $('#lblCounterPartyErrMsg').html(msgText).css('color', '#e81123');
    $('#lblDocumentTypeErrMsg').html(msgText).css('color', '#e81123');
    $('#DocumentTitleErrMsg').html(msgText).css('color', '#e81123');
    $('#lblMasterAgreementNumErrMsg').html(msgText).css('color', '#e81123');

    $('#saveMetadata').hide();
    $('#editMetadata').css('display', 'inline');
    $('#editMetadata').show();
}

//enable metadata fields for editing
var EnableMetadata = function () {
    $("#CounterPartyEdit").removeAttr("disabled");             
    $("#DocumentTypeEdit").removeAttr("disabled");           
    $("#DocumentTitleEdit").removeAttr("disabled");  
    $("#EffectiveDateEdit ").removeAttr("disabled");
    $("#MasterAgreementNumEdit").removeAttr("disabled"); 
    
    $("#CounterPartyEdit").css("background-color", "white");
    $("#DocumentTypeEdit").css("background-color", "white");    
    $("#DocumentTitleEdit").css("background-color", "white");
    $("#EffectiveDateEdit").css("background-color", "white");
    $("#MasterAgreementNumEdit").css("background-color", "white");         

    $('#editMetadata').hide();
    $('#uploadfile').hide();
    $('#message').hide();
    $('#saveMetadata').css('display', 'inline');
    $('#saveMetadata').show();
}

//check if the file is present in Draft library
function checkIfFileExistsInDraft() {
    contractId = $('#ContactId').val();
    $('#progress').hide();

        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('Draft')/Items?$select=Id,Title,ContractID,EffectiveDate,DocumentType,Counterparty,MasterAgreementNumber,CheckoutUserId,OData__UIVersionString&$filter=ContractID eq '"+contractId+"'";
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose"
            },
            success: function (data) {
                if(data.d.results!=null && data.d.results!=undefined && data.d.results.length>0){
                    dataArr =  data.d.results;
                    ItemId = data.d.results[0].ID;
                    
                    $('#search').hide();
                    $('#metadata').show();
                    $('#confirmDetails').show();

                    $("#ContactIdEdit").val(contractId);                
                    $("#CounterPartyEdit").val(data.d.results[0].Counterparty);                
                    $("#DocumentTypeEdit").val(data.d.results[0].DocumentType);                
                    $("#DocumentTitleEdit").val(data.d.results[0].Title);  
                    $("#EffectiveDateEdit ").val(GetFormattedDate(data.d.results[0].EffectiveDate));  
                    $("#MasterAgreementNumEdit").val(data.d.results[0].MasterAgreementNumber); 
                    version = data.d.results[0].OData__UIVersionString;  
                    $("#VersionEdit").val("V "+version); 
                    UpdateItemId = data.d.results[0].ID;

                    $("#ContactIdEdit").css("background-color", "#EBEBE4");
                    $("#CounterPartyEdit").css("background-color", "#EBEBE4");
                    $("#DocumentTypeEdit").css("background-color", "#EBEBE4");    
                    $("#DocumentTitleEdit").css("background-color", "#EBEBE4");
                    $("#EffectiveDateEdit").css("background-color", "#EBEBE4");
                    $("#MasterAgreementNumEdit").css("background-color", "#EBEBE4");
                    $("#VersionEdit").css("background-color", "#EBEBE4");

                    serverRelativUrl = contextURL.split(".com")[1] + "/DRAFT/" + contractId + "/" + contractId + ".docx";
                    console.log('File exists: ' + serverRelativUrl); 
                    getCurrentUser();
                }
                else{
                    console.log('File does not exists in Draft library.'); 
                    var msgText = "This document isn't present in the Draft Library. Please search the Issued Library and retract from approval if you wish to up-revision this document.";
                    $('#msg').html(msgText).css('color', '#e81123');
                    $('#message').show();
                }
                deferred.resolve();
            },
            error: function (error) {
                console.log(error.responseJSON.error.message.value);
                var msgText = "This file is not currently loaded into SharePoint. Please proceed to the manual upload portal to upload this file.";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#message').show();
                deferred.resolve();
            }
        });
        return deferred.promise();

}

//check if the file is present in Issued library
function checkIfFileExistsInIssued() {
    contractId = $('#ContactId').val();
    //alert("Check Issue Library");
    //alert(_spPageContextInfo.webAbsoluteUrl);

        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('Issued')/Items?$select=Id,Title,ContractID,EffectiveDate,DocumentType,Counterparty,MasterAgreementNumber,CheckoutUserId,OData__UIVersionString&$filter=ContractID eq '"+contractId+"'";
        //alert(restUrl);
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose"
            },
            success: function (data) {
                if(data.d.results!=null && data.d.results!=undefined && data.d.results.length>0){
                    //alert("Success");
                    ItemId =  data.d.results[0].ID;
                    $('#search').hide();
                    $('#metadata').show();

                    $("#ContactIdEdit").val(contractId);                
                    $("#CounterPartyEdit").val(data.d.results[0].Counterparty);                
                    $("#DocumentTypeEdit").val(data.d.results[0].DocumentType);                
                    $("#DocumentTitleEdit").val(data.d.results[0].Title); 
                    $("#EffectiveDateEdit ").val(GetFormattedDate(data.d.results[0].EffectiveDate));  
                    $("#MasterAgreementNumEdit").val(data.d.results[0].MasterAgreementNumber);                
                    $("#VersionEdit").val("V "+ data.d.results[0].OData__UIVersionString);   

                    var msgText = "This document is not present in the Draft Library. Please retract form approval if you wish to up-revision this document.";
                    $('#msg').html(msgText).css('color', '#e81123');
                    $('#message').show();
                }
                else{
                    checkIfFileExistsInActive();
                }
                deferred.resolve();
            },
            error: function (error) {
                console.log(error.responseJSON.error.message.value);
                deferred.resolve();
            }
        });
        return deferred.promise();

}

//check if the file is present in Active library
function checkIfFileExistsInActive() {
    contractId = $('#ContactId').val();
    //alert("Check Active Library");
    //alert(_spPageContextInfo.webAbsoluteUrl);

        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('Active')/Items?$select=Id,Title,ContractID,EffectiveDate,DocumentType,Counterparty,MasterAgreementNumber,CheckoutUserId,OData__UIVersionString&$filter=ContractID eq '"+contractId+"'";
        //alert(restUrl);
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose"
            },
            success: function (data) {
                if(data.d.results!=null && data.d.results!=undefined && data.d.results.length>0){
                    //alert("Success");
                    ItemId =  data.d.results[0].ID;
                    $('#search').hide();
                    $('#metadata').show();

                    $("#ContactIdEdit").val(contractId);                
                    $("#CounterPartyEdit").val(data.d.results[0].Counterparty);                
                    $("#DocumentTypeEdit").val(data.d.results[0].DocumentType);                
                    $("#DocumentTitleEdit").val(data.d.results[0].Title);        
                    $("#EffectiveDateEdit ").val(GetFormattedDate(data.d.results[0].EffectiveDate));  
                    $("#MasterAgreementNumEdit").val(data.d.results[0].MasterAgreementNumber);         
                    $("#VersionEdit").val("V "+ data.d.results[0].OData__UIVersionString);

                    var msgText = "This document is not present in the Draft Library. Please retract form approval if you wish to up-revision this document.";
                    $('#msg').html(msgText).css('color', '#e81123');
                    $('#message').show();

                }
                else{
                    if(contractId != "")
                    {
                        var msgText = "Contract with ID " + contractId + " does not exist, please check and try again";
                        $('#msg').html(msgText).css('color', '#e81123');
                        $('#message').show();
                    }
                }
                deferred.resolve();
            },
            error: function (error) {
                console.log(error.responseJSON.error.message.value);
                console.log(error.responseJSON.error.message.value);
                deferred.resolve();
            }
        });
        return deferred.promise();

}


//check if the file is checked out or not
function checkIfThefileIsCheckedOut(UserID) {
    deferred = $.Deferred();
    // alert('inside checking if file is checked out');
    var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + serverRelativUrl + "')/checkOutType";
    // alert(restUrl);
    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json;odata=verbose",
        headers: {
            "Accept": "application/json;odata=verbose"
        },
        success: function (data) {
            checkout = data.d.CheckOutType;
            if (checkout === 0) {
                console.log('The file is checked out');
                GetCheckedOutByUser(UserID);
                deferred.resolve();
            }
            else if (checkout === 2) {
                console.log('The file is not checked out.');
                getCurrentUserPermission();
                $('#buttonmenu').show();
                $('#uploadfile').show();
                deferred.resolve();
            }
        },
        error: function (error) {
            console.log(error.responseJSON.error.message.value);
            serverRelativUrl = contextURL.split(".com")[1] + "/DRAFT/" + contractId + "/" + contractId + ".doc";
            deferred.resolve();
            checkIfThefileIsCheckedOut();
        }
    });
    return deferred.promise();
}

//Get File checkedout by user
function GetCheckedOutByUser(UserID) {
    deferred = $.Deferred();
    var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + serverRelativUrl + "')/CheckedOutByUser";
    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json;odata=verbose",
        headers: {
            "Accept": "application/json;odata=verbose"
        },
        success: function (data) {
           console.log('File is checked out to: ' + JSON.stringify(data.d.Title));
            if(UserID === data.d.Id){
                sameUser = true;
                getCurrentUserPermission();
                $('#buttonmenu').show();
                $('#uploadfile').show();
                deferred.resolve();
            }
            else{
                var msgText = "This document is currently checked out to "+ JSON.stringify(data.d.Title) +". Please check-in the document to proceed.";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#message').show();
            }
        },
        error: function (error) {
            console.log(error.responseJSON.error.message.value);
            deferred.resolve();
        }
    });
    return deferred.promise();
}

function getFileBuffer() {
    var deferred = jQuery.Deferred();
    var fileInput = jQuery('#file');
    var reader = new FileReader();
    reader.onloadend = function (e) {
        deferred.resolve(e.target.result);
    }
    reader.onerror = function (e) {
        deferred.reject(e.target.error);
    }
    reader.readAsArrayBuffer(fileInput[0].files[0]);
    return deferred.promise();
}

//upload file in draft library in the contract folder
function uploadFileSync() {
    return getFormDigest().then(function (data) {
        console.log('inside upload function');
        deferred = $.Deferred();
        var folderUrl = contextURL.split(".com")[1] + "/DRAFT/" + folderName + "/";
        file = $('#file').val();

        var msgText = "Contract upload is in progress.";
        $('#progressMsg').html(msgText).css('color', '#107c10');
        $('#progress').show();

        var getFile = getFileBuffer();
        getFile.done(function (arrayBuffer) {
                var restUrl = _spPageContextInfo.webAbsoluteUrl +
                    "/_api/web/GetFolderByServerRelativeUrl('" + folderUrl + "')/Files/add(url='" + fileName + "', overwrite=true)";
                console.log(restUrl);    
                $.ajax({
                    url: restUrl,
                    type: "POST",
                    data: arrayBuffer,
                    async: false,
                    processData: false,
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue
                        //"content-length": buffer.byteLength
                    },
                    success: function (data) {
                        $('#progress').hide();
                        console.log("Document uploaded successfully. checkin begins");
                        CheckIn();
                        deferred.resolve();
                    },
                    error: function (error) {
                        alert('error');
                        $('#progress').hide();
                        console.log("Upload Failed " + error.responseJSON.error.message.value);
                        var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <click here>";
                        $('#msg').html(msgText).css('color', '#e81123');
                        $('#confirmupload').hide();
                        $('#message').show();
                        deferred.resolve();
                    }
                });
            // }
         });
        return deferred.promise();
    });
}

function pad (number, length) {
    var num = '' + number;
    while (num.length < length) {
        num = '0' + num;
    }
    return num;
  }

//perform chekin after upload
function CheckIn() {
    return getFormDigest().then(function (data) {
        console.log('inside checkin');
        deferred = $.Deferred();
        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + serverRelativUrl + "')/CheckIn(comment='" + $('#comment').val() + "',checkintype=0)";
        $.ajax({
            url: restUrl,
            type: "POST",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue
            },
            success: function (data) {
                console.log("Document was uploaded and checked in successfully");
                var minorVer = version.split('.')[1];
                var len = minorVer.length;
                var minor = "0." + pad("1", len);
                vers = (parseFloat(version) + parseFloat(minor)).toFixed(len);
                if(!sameUser){
                    var msgText = "Document has been uploaded successfully as version "+ vers + "<br>The page will automatically refresh now.";
                }
                else{
                    var msgText = "Document has been uploaded successfully as version "+ version + "<br>The page will automatically refresh now.";
                }
                
                $('#buttonmenu').hide();
                $('#uploadfile').hide();
                $('#confirmupload').hide();
                $('#successMsg').html(msgText).css('color', '#107c10');
                $('#success').show();

                deferred.resolve();
                refreshDocumentLibrary();
            },
            error: function (error) {
                console.log("checkin Failed " + error.responseJSON.error.message.value);
                var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <a href='https://arm.service-now.com'>Click here</a>";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#confirmupload').hide();
                $('#message').show();
                deferred.resolve();
            }
        });
        return deferred.promise();
    });
}

//refresh form digest value
function getFormDigest() {
    return $.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/contextinfo",
        method: "POST",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function () {
            console.log('request digest refreshed');
        },
        error: function (error) {
            console.log("getRequestDigest Failed " + error.responseJSON.error.message.value);
            deferred.resolve();
        }
    });
    return deferred.promise();
}

//perform checkout on the file
function CheckOut() {
    return getFormDigest().then(function (data) {
        deferred = $.Deferred();
        console.log('inside CheckOut function');
        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + serverRelativUrl + "')/CheckOut()";
        // alert(restUrl);
        $.ajax({
            url: restUrl,
            type: "POST",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue
            },
            success: function () {
                console.log("The file has been check-out. Upload process started");
                uploadFileSync();
                deferred.resolve();
            },
            error: function (error) {
                console.log("checkout Failed " + error.responseJSON.error.message.value);
                var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <click here>";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#confirmupload').hide();
                $('#message').show();
                deferred.resolve();
                
            }
        });
        return deferred.promise();
    });
}

function refreshDocumentLibrary() {
    $("#loading").hide();
    setTimeout(function(){ location.reload();}, 3000);
}

function MetaDataCheckOut() {
    return getFormDigest().then(function (data) {
        deferred = $.Deferred();
        console.log('inside metadata checkOut function');
        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + serverRelativUrl + "')/CheckOut()";
        $.ajax({
            url: restUrl,
            type: "POST",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue
            },
            success: function () {
                console.log("The file has been check-out. Metadata update process started");
                SaveData();
                deferred.resolve();
            },
            error: function (error) {
                console.log("checkout Failed " + error.responseJSON.error.message.value);
                var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <click here>";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#confirmupload').hide();
                $('#message').show();
                deferred.resolve();
            }
        });
        return deferred.promise();
    });
}

function SaveData(){
    return getFormDigest().then(function (data) {
        console.log("Save metadata");
        contractId = $('#ContactId').val();
        deferred = $.Deferred();
        var DocType = $("#DocumentTypeEdit").val();  
        var Title = $("#DocumentTitleEdit").val();
        var CounterParty = $("#CounterPartyEdit").val();
        var effectivedate= GetFormattedDate($("#EffectiveDateEdit").val());
        var MasterAgreementNum= $("#MasterAgreementNumEdit").val();
        var VersionNum=$("#VersionEdit").val();
        
        var msgText = "Metadata update is in progress.";
        $('#progressMsg').html(msgText).css('color', '#107c10');
        $('#progress').show();

        var item = {
            "__metadata": { "type": "SP.Data.DraftItem" },
            "DocumentType": DocType,
            "Title": Title,
            "Counterparty": CounterParty,
            "EffectiveDate": GetFormattedDate(effectivedate),
            "MasterAgreementNumber": MasterAgreementNum,
            "OData__UIVersionString": VersionNum
        };                 

        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('Draft')/Items("+UpdateItemId+")";
        $.ajax({
            url: restUrl,
                type: "POST",  
                contentType: "application/json;odata=verbose",  
                data: JSON.stringify(item),  
                headers: {  
                    "Accept": "application/json;odata=verbose",  
                    "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue,  
                    "IF-MATCH": "*",  
                    "X-HTTP-Method":"MERGE",  
                },  
                success: function (data) {
                    $('#progress').hide();
                    console.log('Metadata updated successfully. checkin begins');
                    $("#buttonmenu").hide();
                    MetaDataCheckIn();
                    deferred.resolve();
                },  
                error: function (error) {
                    $('#progress').hide();
                    console.log('Metadata update failed '+ error.responseJSON.error.message.value);
                    var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <click here>";
                    $('#msg').html(msgText).css('color', '#e81123');
                    $('#confirmupload').hide();
                    $('#message').show();
                    deferred.resolve();
                } 
            });
        return deferred.promise();
    });
}

function MetaDataCheckIn() {
    return getFormDigest().then(function (data) {
        console.log('inside metadata checkin');
        deferred = $.Deferred();
        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + serverRelativUrl + "')/CheckIn(comment='Metadata Updated by Commercial-Admin',checkintype=0)";
        $.ajax({
            url: restUrl,
            type: "POST",
            contentType: "application/json;odata=verbose",
            headers: {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": data.d.GetContextWebInformation.FormDigestValue
            },
            success: function (data) {
                console.log("Metadata updated and checked in successfully");
                vers = (parseFloat(version) + parseFloat("0.01")).toFixed(2);
                var msgText = "Metadata of the Contract "+ contractId +" has been uploaded successfully.<br>The page will automatically refresh now.";
                $('#successMsg').html(msgText).css('color', '#107c10');
                $('#buttonmenu').hide();
                $('#confirmupload').hide();
                $('#success').show();
                
                deferred.resolve();
                refreshDocumentLibrary();
            },
            error: function (error) {
                console.log("checkin Failed " + error.responseJSON.error.message.value);
                var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <click here>";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#confirmupload').hide();
                $('#message').show();
                deferred.resolve();
            }
        });
        return deferred.promise();
    });
}

module.exports = {
    OverWriteUpload: OverWriteUpload,
    ConfirmYes: ConfirmYes,
    CheckFileType: CheckFileType,
    ConfirmUpload: ConfirmUpload,
    EnableMetadata: EnableMetadata,
    ReturnToMainMenu: ReturnToMainMenu,
    SaveMetaData: SaveMetaData
};