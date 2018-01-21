// ---------------------------------------------------------------------------------------------------------------------------------------------------
// The MIT License (MIT)
// 
// Copyright (c) 2017 Tom Miller - http://www.miller.mx
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
// ---------------------------------------------------------------------------------------------------------------------------------------------------


function registerWebhook() {
  
  var url = ScriptApp.getService().getUrl();
             
  if (url == null || url == "") {
    
    Browser.msgBox("Please follow instructions on how to publish the script as a web-app:");
      return;
    
  }  
  
  //------------------------------
  
  var error = checkControlValues();
  var monzoUrl = constructMonzoURL("/webhooks");
  
  var headers = {
      "Authorization": "Bearer " + CacheService.getUserCache().get("token")
  };
  
  var data = {
	"account_id": CacheService.getUserCache().get("accountID"),
	"url": url
  };
    
  var options = {
      "headers": headers,
      "method" : "POST",
      'payload' : data,
      "muteHttpExceptions": true
  };

  var resp = UrlFetchApp.fetch(monzoUrl, options);
  
  
  if (resp.getResponseCode() == 200) {
    Browser.msgBox("Webhook successfully registered! PLEASE make sure you change the authorities on the script (See documentation) to allow the webhook callback to work.");
  }  
  else if(resp.getContentText().indexOf("did not return 200 status code, got 403") > 0) {
    Browser.msgBox("Webhook registration failed - HTTP:" + resp.getResponseCode() + ":"
    + " It looks like you need to republish your script with the correct authorities. Please refer to the section in the spreadsheet about generation webhooks. Response from Monzo was: "
    + resp.getContentText());
  }  
  
  else {
    Browser.msgBox("Webhook registration failed - HTTP:" + resp.getResponseCode() + ":" + resp.getContentText());
  }   
}

// This POST is what does all the hard work:   
                                       
function doPost(data) {
  
  var c = data.postData.getDataAsString();
  var action = JSON.parse(c);
 
  logPost(c);
  
  //  We're only interested in actions relating to a card being created on the board.
  if (action.type == "transaction.created") {
    processTransactionFromWebhook(action.data); 
  }  
   
  var x = HtmlService.createHtmlOutput("<p>Roger That</p>")
  return x;
  
}

function logPost(data) {
  var ssId = CacheService.getDocumentCache().get("ssId") || CacheService.getUserCache().get("ssId") || "1ORKw-DOT8fqchgoA5JMM5wpW8KWwd9Csi0TSe3AjFp4"
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName("Logging");
  if (sheet == null) {
    sheet = ss.insertSheet("Logging");
  }  
  sheet.appendRow([new Date(),data]);
  
}  

function processTransactionFromWebhook(data) {
  
  var ssId = CacheService.getDocumentCache().get("ssId") || CacheService.getUserCache().get("ssId") || "1ORKw-DOT8fqchgoA5JMM5wpW8KWwd9Csi0TSe3AjFp4"
  
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName("Transactions");
  sheet.appendRow([getVal(data.id), getVal(data.created), getAmount(data.amount), getVal(data.currency), getAmount(data.local_amount), getVal(data.local_currency), getVal(data.category), getMerchant(data, 'emoji'), getMerchant(data, 'name'), parseAddress(data), getVal(data.notes), getVal(data.receipt)]);
}

// ------------------------------------------------------------------------------------
// Common Functions:
// ------------------------------------------------------------------------------------


function sendError(text) {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Error Updating Transactions" , "The following error occurred processing script to update the transaction list: " + text);
}  

function getAmount(amount) {
  if (amount) {
    return amount/100
  } 
  
  return ""
}

function getMerchant(data, key) {
  if (data.merchant) {
    return data.merchant[key]
  } 
  
  return ""
}

function getVal(val) {
  if (val) {
    return val
  } 
  
  return ""
}


function parseAddress(data) {
  if (data.merchant) {
    if (data.merchant.address) {
      return data.merchant.address.address + ", " + data.merchant.address.city + ", " + data.merchant.address.postcode + ", " + data.merchant.address.country
    }
  }
  return ""
}


function constructMonzoURL(baseURL){
  return "https://api.monzo.com"+ baseURL;
}

function checkControlValues() {
  
  var ssId = CacheService.getDocumentCache().get("ssId") || CacheService.getUserCache().get("ssId")
 
  var col = SpreadsheetApp.openById(ssId).getSheetByName("Control").getRange("B4:B7").getValues();
  var accountID = col[0][0].toString().trim();
  
  if(accountID == "") {
    return "Account Id not found";
  }  
  CacheService.getUserCache().put("accountID", accountID);
  
  var token = col[1][0].toString().trim();
  if(token == "") {
    return "Token not found";
  }  
  CacheService.getUserCache().put("token", token);
  
  CacheService.getUserCache().put("ssId", col[2][0].toString().trim());
  CacheService.getUserCache().put("logPost", col[3][0].toString().trim());
  
  return "";
  
} 

// ------------------------------------------------------------------------------------
// Adding menu options to spreadsheet:
// ------------------------------------------------------------------------------------

function onOpen(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Register Webhook For Transactions", functionName: "registerWebhook"}];
  ss.addMenu("Monzo", menuEntries);

  CacheService.getDocumentCache().put("ssId", ss.getId());
  
  var col = ss.getSheetByName("Control").getRange("B6:B7").getValues();
  
  CacheService.getUserCache().put("logPost", col[0][0].toString().trim());
  CacheService.getUserCache().put("ssId", col[1][0].toString().trim());
 }

