function onInstall(){
  onOpen();
}

function onOpen(e) {

}

function onHomepage(e) {
    //This is for taskbar-menu items. We can explore using those later
    // SpreadsheetApp.getUi()
    //   .createMenu('My Menu')
    //   .addItem('My menu item', 'myFunction')
    //   .addSeparator()
    //   .addSubMenu(SpreadsheetApp.getUi().createMenu('My sub-menu')
    //       .addItem('One sub-menu item', 'mySecondFunction')
    //       .addItem('Another sub-menu item', 'myThirdFunction'))
    //   .addToUi();

  creatonedittrigger('logPush');

  var builder = CardService.newCardBuilder();

  Logger.log("Token valid: " + isTokenValid());

  if(isTokenValid()) {

    builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setText("Build Territory Map")
      .setBottomLabel("Choose the fields that will help you manage your territory.")
      .setOnClickAction(CardService.newAction().setFunctionName('handleTerritory'))
      .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.MULTIPLE_PEOPLE))
      .setWrapText(true)));

    builder.addSection(CardService.newCardSection()
      .setHeader("Actions")
      .addWidget(CardService.newDecoratedText()
        .setText("Configure Thresholds")
        .setOnClickAction(CardService.newAction().setFunctionName('goToThresholdBuilder'))
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION))
        .setBottomLabel("Get notified when specified fields match desired criteria.")
        .setWrapText(true))
      .addWidget(CardService.newDecoratedText()
        .setText("Set Reminder")
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
        .setBottomLabel("Set reminders for yourself related to certain accounts.")
        .setWrapText(true))
    )

  } else {
    builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Authenticate')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOpenLink(CardService.newOpenLink()
            .setUrl(getURLForAuthorization()+"'target='_blank'"))
        .setDisabled(false))));

  }

  return builder.build();
}

//This is an implicit function to Google Apps Scripts (https://developers.google.com/apps-script/guides/html)
function doGet(e) {
  var HTMLToOutput;

  // if(e.parameter.code){//if we get "code" as a parameter in, then this is a callback. we can make this more explicit
  //   // Logger.log("Code: " + e.parameter.code);
    getAndStoreAccessToken2(e.parameters.code);
    HTMLToOutput = '<html><h1>Finished authenticating</h1>Redirecting back to spreadsheet.</html>';
  // }
  return HtmlService.createHtmlOutput(HTMLToOutput);
}

////oAuth related code
function salesforceEntryPoint(URL,method, payload, muteHttpExceptions){
  if(isTokenValid()){

    return SFDChttpRequest(URL, method, payload, muteHttpExceptions);
  }
  else {//we are starting from scratch or resetting
    HTMLToOutput = "<html><h1>You need to login</h1><a href='"+getURLForAuthorization()+"'target='_blank'>Click here to start</a><br>Refresh this page when you return.</html>";
    SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutput(HTMLToOutput));
    return;
  }
}

function SFDChttpRequest(url, requestMethod, payload, muteHttpExceptions){

  var token = userProperties.getProperty(tokenPropertyName);
  var requestDetails = {
    "contentType" : "application/json",
    "method":requestMethod,
    "payload":payload,
    "muteHttpExceptions": muteHttpExceptions,
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json",
    }
  };

  var request = UrlFetchApp.fetch(url,requestDetails);
  // Logger.log("SFDC Response: " + request);
  return request;
}

//hardcoded here for easily tweaking this. should move this to ScriptProperties or better parameterize them
//step 1. we can actually start directly here if that is necessary
var AUTHORIZE_URL = 'https://login.salesforce.com/services/oauth2/authorize'; 
//step 2. after we get the callback, go get token
var TOKEN_URL = 'https://login.salesforce.com/services/oauth2/token'; 

//PUT YOUR OWN SETTINGS HERE
var CLIENT_ID = '3MVG9sn24bYFReCUDUxZgA5NMC6kyJ8qTWFftxmFlN.UtodL3rxmPWh1.WFpaAHf7_rpNNN.0mngxTxWDK2vy';
var CLIENT_SECRET='F8B5FEF0C84D1754AD9BBC478B64383CBF6773EF2950BD04E993FB18421B4245';
var REDIRECT_URL= "https://script.google.com/macros/s/AKfycbxLV7S67snCD-4hyox0dKn4PtA1iPBnYCcwLho90tU/dev" //ScriptApp.getService().getUrl();

//this is the user propety where we'll store the token, make sure this is unique across all user properties across all scripts
var userProperties = PropertiesService.getUserProperties();

var tokenPropertyName = 'SALESFORCE_OAUTH_TOKEN'; 
var baseURLPropertyName = 'SALESFORCE_INSTANCE_URL'; 
var refreshTokenName = 'SALESFORCE_REFRESH_TOKEN';

//this is the URL where they'll authorize with salesforce.com
//may need to add a "scope" param here. like &scope=full for salesforce
function getURLForAuthorization(){
  Logger.log("Authorize URL: " + AUTHORIZE_URL + '?response_type=code&client_id='+CLIENT_ID+'&redirect_uri='+REDIRECT_URL+'&display=page&prompt=select_account') //+'&scope=full')
  return AUTHORIZE_URL + '?response_type=code&client_id='+CLIENT_ID+'&redirect_uri='+REDIRECT_URL+'&display=page&prompt=select_account' //+'&scope=full'
  
}

function getAndStoreAccessToken2(code){

  var nextURL = TOKEN_URL + '?client_id='+CLIENT_ID+'&client_secret='+CLIENT_SECRET+'&grant_type=authorization_code&redirect_uri='+REDIRECT_URL+'&code=' + code+'';
  var options = {
    'method':'post'
  }
  var response = UrlFetchApp.fetch(nextURL, options).getContentText();   
  var tokenResponse = JSON.parse(response);

  //salesforce requires you to call against the instance URL that is against the token (eg. https://na9.salesforce.com/)
  userProperties.setProperty(refreshTokenName, tokenResponse.refresh_token);
  userProperties.setProperty(baseURLPropertyName, tokenResponse.instance_url);
  //store the token for later retrival
  userProperties.setProperty(tokenPropertyName, tokenResponse.access_token);
}

function authTokenGETCheck(muteHttpExceptions) {
  var token = userProperties.getProperty(tokenPropertyName);
  return {
    "contentType" : "application/json",
    "muteHttpExceptions": muteHttpExceptions,
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json",
    }
  };
}

function isTokenValid() {

  if(!userProperties.getProperty(baseURLPropertyName)) {

    Logger.log("Base URL: " + userProperties.getProperty(baseURLPropertyName))

    return false;
  } else {
      var getDataURL = userProperties.getProperty(baseURLPropertyName) + '/services/data/v57.0/query/?q=SELECT+name+from+Account+LIMIT+1';
      var dataResponse = UrlFetchApp.fetch(getDataURL,authTokenGETCheck(true)); 

      if(dataResponse.getResponseCode() === 200) {

        return true;
      } else {
        return false;
      }
  }
}

























