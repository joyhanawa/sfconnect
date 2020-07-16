//PRODUCTION SCRIPT
// JOY HANAWA 2020-07-15
// Use the Analytics API to get report data
// The try-catch can be removed if you do not need to track errors
// This script pulls in the accounts from a report, 
// When you filter from the script, it overrides existing filters.
// The workaround is to pull the existing filters using "describe" and then push the new filter into the object.
// To fetch more than 2000 rows you need to add a sort to your report and then fetch where the value is greater than the last value in your output.
// SF requires a POST action when you use filters, not GET.
// You have to first describe the report, so you can get the column names, they are not what they appear to be, you will need the reporting api name.
// The salesforce report used for this process can be found here: https://yourdomain/lightning/r/Report/{id of your report}/view?queryScope=userFolders
// Run from mainImportClients()  - this will pull the report describe information, needed to build the filter
// Once the filter is built, the sheet is cleared of old data
// The report is fetched and the rows are written to the sheet, the last sort column value is returned as the minimum value for the next run of the report.
// The script checks if there are more rows, if so, it pulls the report again where the sort column is greater than the minimum value.
// This loops until there are no more rows available.


/**** GLOBALS ******/
var REPORT_ID=('***your report id goes here***'); 
var SPREADSHEET_ID = ('***your spreadsheet id goes here***); // name of your sheet
var SHEET_NAME=('*** tab name here ***');



function mainImportClients(){
var newFilters = buildFilters(0); // build new filters to allow mulitple pulls
//Logger.log (newFilters);
clearSheet(REPORT_ID, SHEET_NAME);  
pullReport(REPORT_ID, SHEET_NAME, newFilters,true);  
}

function describeReport(REPORT_ID) {
  var sfService = getSfService();
  sfService.refresh(); 
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var describeUrl = instanceUrl + "/services/data/v47.0/analytics/reports/" + REPORT_ID + "/describe";  // Actual request for report Data
  var response = UrlFetchApp.fetch(describeUrl, { method : "GET", headers : { "Authorization" : "OAuth "+sfService.getAccessToken() } });
  //Logger.log(response);
  var describeResult = JSON.parse(response.getContentText());
  var moreData = describeResult.reportMetadata;
  //Logger.log(moreData);
  var columnNames = describeResult.reportMetadata.detailColumns;
  //Logger.log("columns: \n" + columnNames);  // uncomment this row to see the reporting api column names in the log.
  var reportFilters = describeResult.reportMetadata.reportFilters
  //Logger.log(reportFilters);
  return(reportFilters);
}

function buildFilters(minValue){
  var reportFilters = describeReport(REPORT_ID);
  // This filter repeats the filters in the salesforce report
  // It prepends a filter on my custom field which will allow us to fetch more than 2,000 rows
  // after each POST find the last value and make that minimum for then next report request (POST) 
  var newFilters=reportFilters;
  var temp={};
  var temp = {column: 'Account.Client_Number__c', value : minValue, operator : 'greaterThan'};
  newFilters.push(temp);
  Logger.log("\n number of filters: " + newFilters.length);
  return (newFilters);
  
}

function pullReport(REPORT_ID, SHEET_NAME, newFilters,header) {
  var moreRecords=("TRUE");
  var sfService = getSfService();
  sfService.refresh(); 
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var queryUrl = instanceUrl + "/services/data/v47.0/analytics/reports/" + REPORT_ID + "?includeDetails=true";  // Actual request for report Data
  var payload = JSON.stringify({
            'reportMetadata': {
              'reportFilters': newFilters
             }
      });
  //Logger.log("payload: " + payload)   
  var options = {
            'headers': {
            'Authorization' : 'Bearer ' + sfService.getAccessToken()
            },
            'contentType' : 'application/json',
            'method' : 'POST',
            'payload': payload
  };
 
  var response = UrlFetchApp.fetch(queryUrl, options);
  var queryResult = JSON.parse(response.getContentText());
  var answer = queryResult.factMap["T!T"].rows;  // assumes tabular report  
  var headers = queryResult.reportExtendedMetadata.detailColumnInfo;
  var headname = queryResult.reportMetadata.detailColumns;

  var myArray = [];
  var tempArray = [];
  for (i = 0 ; i < headname.length ; i++) {
    //tempArray.push(headers[headname[i]].label);  //  Use this if you want the column names instead of the API field names
    tempArray.push(headname[i].replace("Account.", ""));  // this is all from the Accounts table.  // We use the API name to update salesforce.
  }
  if(header){
  myArray.push(tempArray);
  // write the header row
  }
  
  
  for (i = 0 ; i < answer.length ; i++ ) {
    var tempArray = [];
    function getData(element,index,array) {
      tempArray.push(array[index].label)
    }
    answer[i].dataCells.forEach(getData);
    myArray.push(tempArray);
  }
    
  var minValue = writeSheet(REPORT_ID, SHEET_NAME, myArray)
 
  if(!queryResult.allData){  // checks for more records, if false, then there is more data and we need to keep pulling the report
    Logger.log("there is more data and the min value is: " + minValue)
      var newFilters = buildFilters(minValue)
      //Logger.log(newFilters);
      pullReport(REPORT_ID, SHEET_NAME, newFilters,false);  // will run the report again, but will not add the header row.
    }else{
     Logger.log("there is NO MORE data")
    };
}


function clearSheet(REPORT_ID, SHEET_NAME){
  var ss= SpreadsheetApp.openById(SPREADSHEET_ID);  // dev Sheet
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1) lastRow = 1;
  sheet.getRange(1,1,lastRow,lastColumn).clearContent();
}


function writeSheet(REPORT_ID, SHEET_NAME, data){
  var ss= SpreadsheetApp.openById(SPREADSHEET_ID);  // dev Sheet
  var sheet = ss.getSheetByName(SHEET_NAME);
  var nextRow = sheet.getLastRow()+1;
  sheet.getRange(nextRow,1, data.length, data[0].length).setValues(data);
  var minValue = sheet.getRange(sheet.getLastRow(),1, 1).getValue();
  return(minValue);
}
