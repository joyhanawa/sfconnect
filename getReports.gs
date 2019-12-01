// Use the Analytics API to get report data
// The try-catch can be removed if you do not need to track errors
function makeRequest(reportId, sheetName) {
  var sfService = getSfService();
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var queryUrl = instanceUrl + "/services/data/v29.0/analytics/reports/" + reportId + "?includeDetails=true";  // Actual request for report Data
  var response = UrlFetchApp.fetch(queryUrl, { method : "GET", headers : { "Authorization" : "OAuth "+sfService.getAccessToken() } });
  var queryResult = JSON.parse(response.getContentText());
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  
  var answer = queryResult.factMap["T!T"].rows;  // assumes tabular report
  var headers = queryResult.reportExtendedMetadata.detailColumnInfo;
  var headname = queryResult.reportMetadata.detailColumns;
  var myArray = [];
  var tempArray = [];
  for (i = 0 ; i < headname.length ; i++) {
    tempArray.push(headers[headname[i]].label);
  }
  myArray.push(tempArray);
  
  for (i = 0 ; i < answer.length ; i++ ) {
    var tempArray = [];
    function getData(element,index,array) {
      tempArray.push(array[index].label)
    }
    answer[i].dataCells.forEach(getData);
    myArray.push(tempArray);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) lastRow = 1;
  sheet.getRange(1,1,lastRow, myArray[0].length).clearContent();
  sheet.getRange(1,1, myArray.length, myArray[0].length).setValues(myArray);
}

// SOQL version of makeRequest()
// The try-catch can be removed if you do not need to track errors
function makeRequestSoql(soql, sheetName) {
  var sfService = getSfService();
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var queryUrl = instanceUrl + "/services/data/v47.0/query?q="+encodeURIComponent(soql);  // Actual request for report Data
  var response = UrlFetchApp.fetch(queryUrl, { method : "GET", headers : { "Authorization" : "OAuth "+sfService.getAccessToken() } });
  var queryResult = JSON.parse(response.getContentText());
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  
  var answer = queryResult.records;  // assumes tabular report
  var fields = soql.substring(7, soql.toUpperCase().indexOf('FROM')-1);
  var headers = fields.split(",");
  var myArray = [headers];
  var tempArray = [];
  
  for (i = 0 ; i < answer.length ; i++ ) {
    for (j = 0 ; j < headers.length ; j++) {
      try {
        var varName="";
        var varName=headers[j].trim(); //some headers have spaces
        var valuePair = answer[i];
        var varValue = valuePair[varName];
        tempArray.push(varValue);
      }
      catch(e) {
        varValue = "";
        break;
      }
    }
    myArray.push(tempArray);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) lastRow = 1;
  sheet.getRange(1,1,lastRow, myArray[0].length).clearContent();
  sheet.getRange(1,1, myArray.length, myArray[0].length).setValues(myArray);
}
