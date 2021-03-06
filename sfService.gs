/*
 * OAuth 2.0 for Salesforce
 * Using OAuth2 for Apps Script - https://github.com/googlesamples/apps-script-oauth2
 * And Salesforce Analytics API - http://www.salesforce.com/us/developer/docs/api_analytics/salesforce_analytics_rest_api.pdf
 
 * Steps to implement: 
 * 1. Create spreadsheet (or document) and paste this code in the associated script project
 * 2. Find the project's key
 * 3. Create a connected app in Salesforce with a callback URL https://script.google.com/macros/d/{SCRIPT ID}/usercallback (this requires Admin access)
 * 4. Get the Consumer Key and Consumer Secret and store them in the Script Properties (in the code below as "sfConsumerKey" and "sfConsumerSecret")
 -  These are specific to the app created.  Currently this app is connected to my sandbox, when it's moved to production you will need new key and secret
 * 5. Set the Project Key and Scope in the function getSfService() 
 * 6. Run the showSidebar() function and click the Authorization Url in the sidebar in your spreadsheet/document (reopen your script project if the sidebar doesn't appear)
 * 7. Login and approve the connected app
 
 * At this point the project has an access token that will expire within a certain time.  If your connected app does not have the "refresh_token" scope, you'll have to 
 * clear the service (clearService()) and repeats steps 6 and 7 in order to get another access token.  
 
 * While your access token is valid you can use it to request reports or make queries. Use makeRequest() and makeRequestSoql() to get data from Salesforce
 
 * If your connected app does have the "refresh_token" scope, use the refreshToken() function to update the access token
 */

// Create the Service
function getSfService() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return OAuth2.createService('salesforce')
    //.setAuthorizationBaseUrl('https://login.salesforce.com/services/oauth2/authorize')  // Production
    //.setTokenUrl('https://login.salesforce.com/services/oauth2/token')  // Production
    .setAuthorizationBaseUrl('https://test.salesforce.com/services/oauth2/authorize')  // sandbox
    .setTokenUrl('https://test.salesforce.com/services/oauth2/token')  // sandbox
    .setClientId(scriptProperties.getProperty("sfConsumerKey"))  // Added in Script Properties
    .setClientSecret(scriptProperties.getProperty("sfConsumerSecret"))  // Added in Script Properties
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('api refresh_token')  // https://help.salesforce.com/HTViewHelpDoc?id=remoteaccess_oauth_scopes.htm&language=en_US
}

// Open a sidebar in the spreadsheet (or document) that will create a link that will take the user authorize the app.
function showSidebar() {
  var sfService = getSfService();
  var test = sfService.hasAccess();
  if (!test) {
    var authorizationUrl = sfService.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);  // If you're using a Document, use DocumentApp instead of SpreadsheetApp
  } else {
    SpreadsheetApp.getActive().toast('Authorization already done.');  // If you're using a Document, use DocumentApp instead
  }
}

// This function is run after the link in the sidebar is clicked and the user authorizes the app. 
function authCallback(request) {
  var sfService = getSfService();
  var isAuthorized = sfService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

// This function clears the service
function clearService() {
  OAuth2.createService('salesforce')
  .setPropertyStore(PropertiesService.getUserProperties())
  .reset();
}
 
