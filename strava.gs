var CLIENT_ID = '<ClientId for the Strava App>';
var CLIENT_SECRET = '<Client Secret for the Strava App>';
var SPREADSHEET_NAME = "StravaData";
var SPREADSHEET_ID = "<Spreadsheet id for the Google Spreadsheet>";
var SHEET_NAME = "Sheet1";
var DEBUG = false;


/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('Strava')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
      .setTokenUrl('https://www.strava.com/oauth/token')

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function that should be invoked to complete
      // the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getService();
  service.reset();
}


/**
 * Authorizes and makes a request to the GitHub API.
 */
function run() {
  var service = getService();
  if (service.hasAccess()) {
    var url = 'https://www.strava.com/api/v3/athlete';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

function retrieveData() {
  //if sheet is empty retrieve all data
  var service = getService();
  if (service.hasAccess()) {
    var sheet = getStravaSheet();
    var unixTime = retrieveLastDate(sheet);
    
    var url = 'https://www.strava.com/api/v3/athlete/activities?after=' + unixTime;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    
    var result = JSON.parse(response.getContentText());

    if (result.length == 0) {
      Logger.log("No new data");
      return;
    }
    
    var data = convertData(result);
    
    if (data.length == 0) {
      Logger.log("No new data with heart rate");
      return;
    }
    
    insertData(sheet, data);
   
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

function retrieveLastDate(sheet) {
  var lastRow = sheet.getLastRow();
  var unixTime = 0; 
  if (lastRow > 0) { 
      var dateCell = sheet.getRange(lastRow, 1);
      var dateString = dateCell.getValue();
      var date = new Date((dateString || "").replace(/-/g,"/").replace(/[TZ]/g," "));
      unixTime = date/1000;
   }
   return unixTime;
}

function convertData(result) {
  var data = [];
  
  for (var i = 0; i < result.length; i++) {
    if (result[i]["has_heartrate"]) {
      var item = [result[i]['start_date_local'],
                  result[i]['max_heartrate'],
                  result[i]['average_heartrate']];
      data.push(item);
    }
      
  }
  
  return data;
}

function getStravaSheet() {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);
  return sheet;
}

function insertData(sheet, data) {
  var header = ["Date", "MaxHeartRate", "AvgHeartRate"];
  ensureHeader(header, sheet);
  
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow+1,1,data.length,3);
  range.setValues(data); 
}

function ensureHeader(header, sheet) {
  // Only add the header if sheet is empty
  if (sheet.getLastRow() == 0) {
    if (DEBUG) Logger.log('Sheet is empty, adding header.')    
    sheet.appendRow(header);
    return true;
    
  } else {
    if (DEBUG) Logger.log('Sheet is not empty, not adding header.')
    return false;
  }
}


function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    if (DEBUG) Logger.log('Sheet "%s" does not exists, adding new one.', sheetName);
    sheet = spreadsheet.insertSheet(sheetName)
  } 
  
  return sheet;
}