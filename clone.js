
//************************************ Storing credentials securely ***********************************************//
PropertiesService.getScriptProperties().setProperty('clientId', '3MVG9kBt168mda_.FRkRWH9jKEI1Gc4pvG8Xj84GeHcQ6V8Qh_i0G8kFqNchRxvaS0k8gx6YYqhx1o7Pc8iLA');
PropertiesService.getScriptProperties().setProperty('clientSecret', 'B566BBF8643EC230DD8E2B8A1C409AB14624701101C1D1C17D65F4012FA8A7A2');
PropertiesService.getScriptProperties().setProperty('redirectUri', 'https://script.google.com/macros/d/1pmloc0p_RFF8RzSyW1d2xKvY2z8NolVOMdP-KT6J2syrwwyxH7Z5r6Wx/usercallback');
PropertiesService.getScriptProperties().setProperty('scope', 'refreshToken, api');

// -----------------------------------------------------------------------------------------------------------------//
// ------------------------------------Methods-For-Org Authorization------------------------------------------------//
// -----------------------------------------------------------------------------------------------------------------//

let siteBaseUrl;

//************************************ Authorization Starts From Here *********************************************//
function startOAuth2WithAuthUrl(authUrl) {

  clearStoredAccessTokenAndInstanceUrl(); // clear previously saved user properties
  var selectedAuthUrl = authUrl;
  let authorizationUrl;
  var accessToken;
  var instanceUrl

  let clientId = PropertiesService.getScriptProperties().getProperty('clientId');
  let clientSecret = PropertiesService.getScriptProperties().getProperty('clientSecret');
  let redirectUri = PropertiesService.getScriptProperties().getProperty('redirectUri');
  let scope = PropertiesService.getScriptProperties().getProperty('scope');


  var service = OAuth2.createService('InterACTdataLoaderAuthorization')
    .setAuthorizationBaseUrl(selectedAuthUrl + '/services/oauth2/authorize')
    .setTokenUrl(selectedAuthUrl + '/services/oauth2/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())

  authorizationUrl = service.getAuthorizationUrl() + '&prompt=login';

  Logger.log(authorizationUrl);

  return authorizationUrl;
}

//************************************ Auth Call Back Function ***************************************************//
function authCallback(request) {

  let clientId = PropertiesService.getScriptProperties().getProperty('clientId');
  let clientSecret = PropertiesService.getScriptProperties().getProperty('clientSecret');
  let redirectUri = PropertiesService.getScriptProperties().getProperty('redirectUri');
  let scope = PropertiesService.getScriptProperties().getProperty('scope');

  var service = OAuth2.createService('InterACTdataLoaderAuthorization')
    .setAuthorizationBaseUrl('https://login.salesforce.com/services/oauth2/authorize')
    .setTokenUrl('https://login.salesforce.com/services/oauth2/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback') 
    .setPropertyStore(PropertiesService.getUserProperties()) 

  var authorized = service.handleCallback(request);

  if (authorized) {
    var accessToken = service.getAccessToken();
    var instanceUrl = getInstanceUrl(accessToken);
    var siteBaseUrl;

    if (service && service.hasAccess()) {

      var apiUrl = instanceUrl + "/services/data/v53.0/query?q=SELECT+Id,Domain.Domain,Site.Name+FROM+DomainSite";
      var headers = {
        Authorization: "Bearer " + accessToken
      };

      var options = {
        method: "get",
        headers: headers
      };

      var response = UrlFetchApp.fetch(apiUrl, options);
      var responseData = JSON.parse(response.getContentText());

      siteBaseUrl = (responseData.records.find(record => (record.Site && (record.Site.Name === 'CRMA_InterACTDataLoader' || record.Site.Name === 'InterACTDataloader'))) || {}).Domain?.Domain || '';
    }

    // Store the access token and instance URL
    storeAccessTokenAndInstanceUrl(accessToken, instanceUrl, siteBaseUrl);
    invokeHome();


    var successHtml = '<html><head><title>Authorization Successful</title></head><body>'
      + '<h1>Authorization successful!.....</h1>'
      + '<p>You can now close this tab and return to the Google Sheets add-on.</p>'
      + '<script>setTimeout(function() { window.close(); }, 2000);</script>' // Auto-close after 2 seconds
      + '</body></html>';
    var htmlOutput = HtmlService.createHtmlOutput(successHtml);
    return htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    var failureHtml = '<html><head><title>Authorization Failed</title></head><body>'
      + '<h1>Authorization failed!</h1>'
      + '<p>Please try again.</p>'
      + '</body></html>';
    var htmlOutput = HtmlService.createHtmlOutput(failureHtml);
    return htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

//************************************ To clear stored access token, Base Url and instance url *********************//
function clearStoredAccessTokenAndInstanceUrl() {

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('AccessToken');
  scriptProperties.deleteProperty('InstanceUrl');
  scriptProperties.deleteProperty('SiteBaseUrl');

}

//************************************ To store access token, Base Url and instance url ****************************//
function storeAccessTokenAndInstanceUrl(accessToken, instanceUrl, siteBaseUrl) {

  clearStoredAccessTokenAndInstanceUrl();
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('AccessToken', accessToken);
  scriptProperties.setProperty('InstanceUrl', instanceUrl);
  scriptProperties.setProperty('SiteBaseUrl', siteBaseUrl);
}

//************************************ To Get Stored Access Token *************************************************//
function getStoredAccessToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('AccessToken');
}

//************************************ To Get Stored Base Url *****************************************************//
function getStoredBaseUrl() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('SiteBaseUrl');
}

//************************************ To Get Stored Instance Url *************************************************//
function getStoredInstanceUrl() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('InstanceUrl');
}

//************************************ To Get Instance Url *******************************************************//
function getInstanceUrl(accessToken) {

  var response = UrlFetchApp.fetch('https://login.salesforce.com/services/oauth2/userinfo', {
    headers: {
      Authorization: 'Bearer ' + accessToken
    }
  });

  var json = response.getContentText();
  var data = JSON.parse(json);

  if (data && data.urls && data.urls.custom_domain) {
    return data.urls.custom_domain;
  } else if (data && data.urls && data.urls.enterprise) {
    return data.urls.enterprise;
  } else {
    return 'login.salesforce.com';
  }
}



/**
 * Open a URL in a new tab.
 */
function openUrl(url) {
  var html = HtmlService.createHtmlOutput('<html><script>'
    + 'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
    + 'var a = document.createElement("a"); a.href="' + url + '"; a.target="_blank";'
    + 'if(document.createEvent){'
    + '  var event=document.createEvent("MouseEvents");'
    + '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
    + '  event.initEvent("click",true,true); a.dispatchEvent(event);'
    + '}else{ a.click() }'
    + 'close();'
    + '</script>'
    // Offer URL as clickable link in case above code fails.
    + '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
    + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
    + '</html>')
    .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
}


//***********************setCredentials****************************************************//
function getCredentials() {

  const consumerkey = '3MVG9gtDqcOkH4PI9ckehInuTTHcj9X8CpAEAo6fh..ycN1INaa5rBcd0YPEoll_pEg.AVoKYvbLS5Pz48V1f';
  const consumersecret = 'D9AEEFEBD7FEF097A2C7470C2A97DF2AEF09AC22803F71D768FDD70306E71C73';
  const username = 'aliya.ali-9337632919@industryapps.com';
  const password = 'CRMantra@123';
  const securitytoken = 'fVI3ccRMVASO7YSaok1qMhw9';
  return { username: username, password: password, securitytoken: securitytoken, consumerkey: consumerkey, consumersecret: consumersecret }
}

//***********************Create Access Token****************************************************//
function getAccessToken() {
  return { accessToken: getStoredAccessToken(), instanceURL: getStoredInstanceUrl() }
}

//***********************Get Salesforce Objects****************************************************//
function getSFObjects() {
  authResult = getAccessToken()
  var queryUrl = authResult.instanceURL + "/services/data/v55.0/sobjects";
  var header = {
    Authorization: " Bearer " + authResult.accessToken,
  };
  var options = {
    'method': 'get',
    'headers': header
  };
  response = UrlFetchApp.fetch(queryUrl, options);
  var json_response = JSON.parse(response);
  console.log("responseVSP-->" + JSON.stringify(json_response.sobjects));
  Logger.log("responseVSP-->" + JSON.stringify(json_response.sobjects));
  return json_response;
}
