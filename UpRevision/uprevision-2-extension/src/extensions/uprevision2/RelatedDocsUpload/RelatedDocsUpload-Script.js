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
var RelatedDocsUpload = function () {
    $('#metadata').hide();
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
        checkIfContractExistsInActive();
    }
}

//upload document to SharePoint
var ConfirmUpload = function () {
    if($('#file').val().length == 0){
        var msgText = "Please select the file to upload.";
        $('#lblFileMsg').html(msgText).css('color', '#e81123');
    }
    else if($('#comment').val().length == 0){
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
        uploadFileSync();        
    }
}

//click of Return to main menu
var ReturnToMainMenu = function () {
    $('#ContactId').val("");
    $('#search').show();
    $('#metadata').hide();
    $('#message').hide();
    $('#success').hide();
}

//check if the Contract is present in Draft library
function checkIfContractExistsInActive() {
    contractId = $('#ContactId').val();
    $('#progress').hide();
    var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('Active')/Items?$select=Id,ContractID,DocumentType,Counterparty,MasterAgreementNumber&$filter=ContractID eq '"+contractId+"'";
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
                    
                $("#ContactIdEdit").val(contractId);                
                $("#CounterPartyEdit").val(data.d.results[0].Counterparty);                
                $("#DocumentTypeEdit").val(data.d.results[0].DocumentType);                
                $("#MasterAgreementNumEdit").val(data.d.results[0].MasterAgreementNumber); 
                $("#ContactIdEdit").css("background-color", "#EBEBE4");
                $("#CounterPartyEdit").css("background-color", "#EBEBE4");
                $("#DocumentTypeEdit").css("background-color", "#EBEBE4");    
                $("#MasterAgreementNumEdit").css("background-color", "#EBEBE4");
                  
                serverRelativUrl = contextURL.split(".com")[1] + "/ACTIVE/" + contractId + "/" + contractId + ".docx";
                console.log('Contracts exists: ' + serverRelativUrl); 
            }
            else{
                console.log('Contract does not exists in Active library.'); 
                var msgText = "This Contract ID does not exit in Active Library.Please enter the valid Contract ID.";
                $('#msg').html(msgText).css('color', '#e81123');
                $('#message').show();
            }
            deferred.resolve();
        },
        error: function (error) {
            console.log(error.responseJSON.error.message.value);
            var msgText = "This file is not currently loaded into SharePoint. Please try again.";
            $('#msg').html(msgText).css('color', '#e81123');
            $('#message').show();
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
        var folderUrl = contextURL.split(".com")[1] + "/Contract Related Documents/";
        file = $('#file').val();
        fileName = file.substring(file.lastIndexOf('\\') + 1);
        var msgText = "Contract upload is in progress.";
        $('#progressMsg').html(msgText).css('color', '#107c10');
        $('#progress').show();

        var getFile = getFileBuffer();
        getFile.done(function (arrayBuffer) {
                var restUrl = _spPageContextInfo.webAbsoluteUrl +
                    "/_api/web/GetFolderByServerRelativeUrl('" + folderUrl + "')/Files/add(url='" + fileName + "', overwrite=true)?$expand=ListItemAllFields";
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
                        console.log("Document uploaded successfully. Metadata Update Begins");
                        UpdateItemId = data.d["ListItemAllFields"]["ID"];
                        SaveMetaData();
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

//Save Metadata
function SaveMetaData(){
    //alert("Save MetaData");
    return getFormDigest().then(function (data) {
        console.log("Save metadata is in progress");
        contractId = $('#ContactIdEdit').val();
        deferred = $.Deferred();
        var DocType = $("#DocumentTypeEdit").val();  
        var CounterParty = $("#CounterPartyEdit").val();
        var MasterAgreementNum= $("#MasterAgreementNumEdit").val();
        var CommentCheckIn = $("#comment").val();

        var msgText = "Metadata update is in progress.";
        $('#progressMsg').html(msgText).css('color', '#107c10');
        $('#progress').show();

        var item = {
            "__metadata": { "type": "SP.Data.Contract_x0020_Related_x0020_DocumentsItem" },
            "ContractID": contractId,
            "DocumentType": DocType,
            "Counterparty": CounterParty,
            "MasterAgreementNumber": MasterAgreementNum,
            //"comment": CommentCheckIn
        };                 

        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/Web/lists/GetByTitle('Contract Related Documents')/Items("+UpdateItemId+")";
        
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
                    var msgText = "Document has been uploaded successfully. The page will automatically refresh now.";
                    $('#successMsg').html(msgText).css('color', '#107c10');
                    $('#success').show();

                    deferred.resolve();
                    refreshDocumentLibrary();
                },  
                error: function (error) {
                    $('#progress').hide();
                    console.log('Metadata update failed '+ error.responseJSON.error.message.value);
                    var msgText = "Sorry there's been a problem and your document hasn't been uploaded.<br>Please contact the help desk <click here>";
                    $('#msg').html(msgText).css('color', '#e81123');
                    $('#message').show();
                    deferred.resolve();
                } 
            });
        return deferred.promise();
    });
}

//perform chekin after upload
function CheckIn() {
    return getFormDigest().then(function (data) {
        console.log('inside checkin');
        deferred = $.Deferred();
        var folderUrl = contextURL.split(".com")[1] + "/Contract Related Documents/";
        var fileurl = folderUrl + "/" + fileName;
        var restUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + fileurl + "')/CheckIn(comment='" + $('#comment').val() + "',checkintype=0)";
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

module.exports = {
    RelatedDocsUpload: RelatedDocsUpload,
    ConfirmUpload: ConfirmUpload,
    ReturnToMainMenu: ReturnToMainMenu
};