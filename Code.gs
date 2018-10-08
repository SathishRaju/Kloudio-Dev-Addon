var _ = Underscore.load();
var currentQuery = {};
var currentTemplate = null;
var selectedCols = [];
var refreshMode = false;
var currentDataTable = null;

 // var apiUrl = 'https://api-preprod.kloud.io/v1';
 var apiUrl = 'https://api-dev.kloud.io/v1';
 // var apiUrl = 'https://kloudio.ngrok.io/v1';
// var fnUrl = 'http://runner.kloud.io/v1';

 // var appUrl = 'https://app-preprod.kloud.io';
 var appUrl = 'https://app-dev.kloud.io';
// var appUrl = 'http://localhost:8080';

//var apiUrl = 'https://yheqjqsmgs.localtunnel.me/v1';
//var appUrl = 'http://localhost:8080';

var uploadMode = 'All';

JsLinq.GenerateEnumerable(this);
var Enumerable = this.Enumerable;

var batchSize = 5000;

///CHANGE THE VERSION//
var version = '2.0.115';

///CHANGE THE VERSION//

function getKloudioVersion()
{
  SpreadsheetApp.getUi().alert('You are using Kloudio Addon Version: ' + version); 
}


function UnAuthorizedException_(message) {
   this.message = message;
   this.name = 'UnAuthorizedException';
}

function PlanInactiveException_(message) {
   this.message = message;
   this.name = 'PlanInactiveException';
}


Object.defineProperty(Array.prototype, 'chunk', {
    value: function(chunkSize) {
        var R = [];
        for (var i=0; i<this.length; i+=chunkSize)
            R.push(this.slice(i,i+chunkSize));
        return R;
    }
});

function endsWith_(str, suffix) {
  //console.log('endsWith_...');
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
}

function startsWith_(str, prefix) {
    return str.indexOf(prefix) === 0;
}


function showAlert(str)
{
  SpreadsheetApp.getUi().alert(str);
}


/**
* example
* @param {string} input the example URL.
* @return The example URL.
* @customfunction
*/
function GETID(URL) {
    return 'example '+URL;
}


/**
 * Returns amount multiplied by itself.
 *
 * @param {Number} amount The amount to be multiplied by itself.
 * @return {Number} The amount multiplied by itself.
 * @customfunction
 */
function test(amount) {
  return amount*amount;
}


/**
 * Run a Kloud Function.
 *
 * @param {string} Kloud Function Name.
 * @param {string} Parameter1 Name of Parameter1
 * @param {string} Value1 Value of Parameter1
 * @param {string} Parameter2 Name of Parameter2
 * @param {string} Value2 Value of Parameter2
 * @param {string} Parameter3 Name of Parameter3
 * @param {string} Value3 Value of Parameter3
 * @param {string} Parameter4 Name of Parameter4
 * @param {string} Value4 Value of Parameter4 
 * @param {string} Parameter5 Name of Parameter5
 * @param {string} Value4 Value of Parameter5 
 * @return  Data
 * @customfunction
 */
function Kfx(fnName, parameter1, value1, parameter2, value2, parameter3, value3, parameter4, value4, parameter5, value5)
{
  try
  {
    
    console.log("Kloud Function...");
   
    
    var parameters = {};
    
    if(parameter1 != undefined)
      parameters[parameter1] = value1;
    
    if(parameter2 != undefined)
      parameters[parameter1] = value1;
    
    if(parameter3 != undefined)
      parameters[parameter1] = value1;  
    
    if(parameter4 != undefined)
      parameters[parameter1] = value1;   

    if(parameter5 != undefined)
      parameters[parameter1] = value1; 
    
    var newData = [];
    var data = [];

    var response =  executeFunction_(fnName, parameters);  
    if(response.error) {
      console.log("response failed");
      return [decodeResponse_(response)];
    } 
    else 
    { 
      return response.output;
    }
  }
  catch(ex)
  {
    console.log(ex.message)
  }
}



function decodeResponse_(response) {
  if(response.message) {
    return response.message;
  } else {
    if(response.error) {
      if(response.error === 'jwt expired') {
        return "Session Expired";
      } else {
        return response.error;
      }
    } else {
      return "Error";
    }
  }
}
/*function Kloud(reportId, parameter1, value1, parameter2, value2, parameter3, value3, parameter4, value4, parameter5, value5)
{
  try
  {
    
    console.log("Kloud Function...");
    var query = getQueryById_(reportId);
    if(query === null) {
      console.log('query === null');
      return ["UnAuthorized"];
    }
    if(query === -1){
      console.log('query === -1');
      return ["Report not found"];
    }
    console.log("query name = " + query.queryName);
    
    var parameters = [];
    
    if(parameter1 != undefined)
      parameters.push({param_name: parameter1, param_val:value1});
    
    if(parameter2 != undefined)
      parameters.push({param_name: parameter2, param_val:value2});
    
    if(parameter3 != undefined)
      parameters.push({param_name: parameter3, param_val:value3});    
    
    if(parameter4 != undefined)
      parameters.push({param_name: parameter4, param_val:value4});   

    if(parameter5 != undefined)
      parameters.push({param_name: parameter5, param_val:value5}); 
    
    var newData = [];
    var data = [];

    var response =  getDataPvt_(query, parameters,true);  
    if(!response.success) {
      return [response.message];
    } 
    else 
    { 
      return response.data;
    }
    
   
  }
  catch(ex)
  {
    console.log(ex.message)
  }
}

*/


function getApiUrl_()
{
  console.log('getApiUrl_');
  return apiUrl;
}

function getFnUrl_()
{
  console.log('getFnUrl_');
  return fnUrl;
}

function getAppUrl_()
{
  console.log('getApiUrl_');
  var email = Session.getEffectiveUser().getEmail();
  console.log('email = ' + email);
  if(endsWith_(email,'@netflix.com'))
  {
    return 'https://kloudio.itp.netflix.net';
  }
  else
  {
     if(endsWith_(email,'@transenterix.com'))
    {
      return 'http://kloudio.transenterix.com:8380';
    }
    else
      return appUrl;
  }

}

function getAppTestUrl_()
{
  console.log('getAppTestUrl_');
  var email = Session.getEffectiveUser().getEmail();
  console.log('email = ' + email);
  if(endsWith_(email,'@netflix.com'))
  {
     if(_.contains(['sraju@netflix.com', 'rbafna@netflix.com', 'kbhatvendor@netflix.com','ssurapaneni@netflix.com'], email))
      return 'https://kloudio.itd.netflix.net';
    else
      return '';
  }
  else
  {
    return appUrl;
  }

}

function getParseOptions_() {
   var headers = {
    "Content-Type": "application/json",
     "Authorization": "Bearer " + getToken_()
  };
  var options =
      {
        "method" : "get",
        "headers": headers,
        "muteHttpExceptions": true
      };
  return options;
}

function addMenuEntries_() {
  console.log("adding menu entries...");
  
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  if(getToken_() != null)
  {
    console.log("already loggedin...");
    ui.createMenu('Kloudio')
    .addItem('Run Report', 'showReportsPage')
   // .addItem('Refresh Data', 'openParamsDialog')
          .addSubMenu(ui.createMenu('Refresh')
          .addItem('Current Sheet', 'openParamsDialog')
          .addItem('All Sheets','refreshAllReports'))
    .addItem('Refresh Formulas', 'refreshFormulas')
    .addItem('Update Schedule', 'openScheduleDialog')
    .addSeparator()
    .addItem('My Templates', 'showTemplatePanel')
      .addSubMenu(ui.createMenu('Upload')
          .addItem('All', 'uploadData')
          .addItem('Selected','uploadSelectedData'))
    .addSeparator()
    .addItem('Get Trace ID', 'getTraceId')
    .addItem('About', 'getKloudioVersion')
    .addItem('Logout', 'logout')
    .addToUi();
  }
  else
  {
    console.log("not logged in...");
    ui.createMenu('Kloudio')
    .addItem('Login', 'showLoginPage')
    .addToUi();
  }    
}

function logout() {
  try
  {
    console.log("logout");
    var userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty("AUTH_TOKEN");
    userProperties.deleteProperty("INSTANCE_TYPE");
    userProperties.deleteProperty("CURRENT_USER");
    addMenuEntries_();    
    SpreadsheetApp.getActiveSpreadsheet().toast("You have succesfully logged out","Kloudio");
    showLoginPage();
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}

function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Kloudio')
    .addItem('Run Report', 'showReportsPage')
    //.addItem('Refresh Data', 'openParamsDialog')
          .addSubMenu(ui.createMenu('Refresh')
          .addItem('Current Sheet', 'openParamsDialog')
          .addItem('All Sheets','refreshAllReports'))  
    .addItem('Refresh Formulas', 'refreshFormulas')
    .addItem('Update Schedule', 'openScheduleDialog')
//    .addItem('Current Reports', 'showReportList')
    .addSeparator()
    .addItem('My Templates', 'showTemplatePanel')
      .addSubMenu(ui.createMenu('Upload')
          .addItem('All', 'uploadData')
          .addItem('Selected','uploadSelectedData')) 
    .addSeparator()
    .addItem('Get Trace ID', 'getTraceId')
    .addItem('About', 'getKloudioVersion')
    .addItem('Logout', 'logout')
    .addToUi();
  
  // addMenuEntries_();
 
}


function getTraceId()
{
  var sessionId = Session.getTemporaryActiveUserKey();
  SpreadsheetApp.getUi().alert('Kloudio Trace ID: \n' + sessionId + '\n\nPlease provide this ID to the support when reporting any issues.');
}

function spreadOnEdit(e) {
  //console.log('onEdit..');
  //console.log(e);
}


function onInstall(e) {
    addMenuEntries_();
    showLoginPage();
}


function authcallback(request)
{
  console.log('authCallback');
  if(!!!request.parameter.token)
  {
    /** not authorized **/
  }
  else
  {
    //SpreadsheetApp.getUi().alert("got token: " + request.parameter.token);
     var userProperties = PropertiesService.getUserProperties();
     userProperties.setProperty('AUTH_TOKEN', request.parameter.token);
//    console.log('token = ' + request.parameter.token);
    userProperties.setProperty('INSTANCE_TYPE', request.parameter.instance == undefined ? '' : request.parameter.instance);
     addMenuEntries_();
   // showAlert("going to show reports page...");
     showReportsPage();
    return HtmlService.createHtmlOutput('You have been successfully authorized by Kloudio. You may now close this window and start using Kloudio!.');
  }
}


function getInstanceType_()
{
  console.log('getInstanceType_..');
   var userProperties = PropertiesService.getUserProperties();
   var val = userProperties.getProperty("INSTANCE_TYPE");
  console.log('val = ' + val);
  return val;
}

function getToken_()
{
   var userProperties = PropertiesService.getUserProperties();
   var cachedToken = userProperties.getProperty("AUTH_TOKEN");
   if (cachedToken != null) {
     return cachedToken;
   }
  return null;
}

  function getDifferenceInDays_(firstDate , secondDate) {  
      var firstTime = firstDate.getTime();
      var secondTime = secondDate.getTime();
      console.log("firstTime = " + firstTime + ", secondTime = " + secondTime);
        return Math.round((firstDate.getTime() - secondDate.getTime())/(24*60*60*1000));
   }

  function getRemainingDays_(plan_end_date) {
    var numbers = plan_end_date.match(/\d+/g); 
    var first_date = new Date(numbers[0], numbers[1]-1, numbers[2]);
    console.log('plan-end-date = ' + first_date);
        return getDifferenceInDays_(first_date,new Date()) + 1;
    }


function getCallbackURL_(callbackFunction,instanceType) {
  console.log('getCallbackURL_...');
   var scriptUrl = 'https://script.google.com/a/macros/kloud.io/d/' + ScriptApp.getScriptId();
   var urlSuffix = '/usercallback?state=';
   var stateToken = ScriptApp.newStateToken()
       .withMethod(callbackFunction)
       .withTimeout(300)
       .createToken();
  var url =  scriptUrl + urlSuffix + stateToken + (instanceType != undefined ? '&instance=' + instanceType : '') ;
  console.log('url = ' + url);
  return url;
 }


function getTrialEndingText_() {  
  console.log("getTrialEndingText_...");
  
  var user = currentUser_();
  if(user.planName == 'Trial')
  {
  console.log("user's trial ending date = " + user.planEndDate);
  var remainingDays = getRemainingDays_(user.planEndDate);
  console.log("remainingDays = " + remainingDays);
  if(remainingDays < 0){
    return {text: 'YOUR TRIAL HAS ALREADY ENDED', bgColor: 'uk-alert-danger'};
  }
  else
  { 
    if(remainingDays == 0)
      return {text: 'YOUR TRIAL ENDS TODAY', bgColor: 'uk-alert-danger'};
    if(remainingDays == 1)
      return {text: 'YOUR TRIAL ENDS TOMORROW', bgColor: 'uk-alert-danger'};
    
    return  {text: 'YOUR TRIAL ENDS IN ' + remainingDays + ' DAYS' , bgColor: remainingDays > 6 ? 'uk-alert-primary' : 'uk-alert-danger'};
  }
  }
  return {text: '', bgColor: ''};
}

function getPlanCanceledText_(){
  console.log("getPlanCanceledText_...");
  var user = currentUser_();
  if(user.planName == 'Pro' || user.planName == 'Enterprise' || user.planName == 'Premium')
  {
      return 'YOUR PLAN IS CANCELED';
  }
  else
    return '';
}


function checkTrialEnded_()
{
  var user = currentUser_();
  if(user)
  {
    return user.planName == 'Trial' && getRemainingDays_(user.planEndDate) <=0;
  }
  else
    return true;
}

function checkPlanCanceled_()
{
  console.log("checkPlanCanceled_...");
  var user = currentUser_();
  if(user)
  {
    return (user.planName == 'Pro' || user.planName == 'Enterprise' || user.planName == 'Premium') && ((user.planStatus == 'Canceled' && getRemainingDays_(user.planEndDate) <= 0) || user.planStatus == 'Suspended');
  }
  else
    return true;
}

function checkForExpiration_()
{
  //showAlert("checkForExpiration_...");
   console.log('checkForExpiration_...');
  
   var user = getCurrentUser_();
  //showAlert("after getcurrentuser...");
   var email = Session.getEffectiveUser().getEmail();
   if(!endsWith_(email,'@netflix.com')) {     
     return checkTrialEnded_() || checkPlanCanceled_();
   }
  else
    return false;
}

function generateLoginPage_()
{
  console.log('generateLoginPage_...');
 
  var email = Session.getEffectiveUser().getEmail();
  
    var template;
    var page;
    if(endsWith_(email,'@netflix.com'))
    {
      template = HtmlService.createTemplateFromFile('login-nflx');
      template.authorizationProdUrl = getAppUrl_() + '/gsheets?src=gsheets&redirectUrl=' + encodeURIComponent(getCallbackURL_('authcallback','prod'));
      template.authorizationTestUrl = getAppTestUrl_() + '/gsheets?src=gsheets&redirectUrl=' + encodeURIComponent(getCallbackURL_('authcallback','tst'));
    }
    else
    {
      template = HtmlService.createTemplateFromFile('login');
      template.authorizationUrl = getAppUrl_() + '#/user/login?src=gsheets&redirectUrl=' + encodeURIComponent(getCallbackURL_('authcallback'));
    }
  
    page = template.evaluate();
    page.setTitle("Kloudio Login")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return page;
}

function generateQueriesPage_()
{
  try{
  console.log("generatingQueriesPage()");
   // showAlert("generateQueriesPage_");
  var email = Session.getEffectiveUser().getEmail();
  var template = HtmlService.createTemplateFromFile('index');
  if(endsWith_(email,'@netflix.com'))
  {
    if(getInstanceType_() == 'tst')
    {
      template.appUrl = getAppTestUrl_();
    }
    else
    {
      template.appUrl = getAppUrl_();
    }
  }
  else
  {
    template.appUrl = getAppUrl_();
  }
  if(checkForExpiration_()){
    showUpgradePage_();
    return null;   
  }
  else
  {
    console.log('Plan is active...');
    var user = currentUser_();
    var trialEnding = getTrialEndingText_();
    template.trialText = user.planName == 'Trial' ? trialEnding.text : '';
    template.upgradeText = user.planName == 'Trial' ? 'UPGRADE NOW' : '';
    template.bgColor = trialEnding.bgColor;
    template.upgradeUrl = getAppUrl_() + '/#/app/account/upgrade';
  }
    var page = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Kloudio - Run Report')
      .setWidth(300);
  console.log("after generating page...");
   console.log('Access token = ' + ScriptApp.getOAuthToken());
  return page;
  }
  catch(ex)
  {
    console.log(ex);
    if(ex.name == 'UnAuthorizedException')
      handleUnAuthorized_();
     if(ex.name == 'PlanInactiveException')
       showUpgradePage_();
  }
}

function generateTemplatesPage_()
{
  console.log("generateTemplatesPage_()");
  var email = Session.getEffectiveUser().getEmail();
  var template = HtmlService.createTemplateFromFile('Templates');
  if(endsWith_(email,'@netflix.com'))
  {
    template.appUrl = getAppTestUrl_();
  }
  else
  {
    template.appUrl = getAppUrl_();
  }
  
  if(checkForExpiration_()){
    showUpgradePage_();
    return null;
  }
  else
  {
    var user = currentUser_();
    var trialEnding = getTrialEndingText_();
    template.trialText = user.planName == 'Trial' ? trialEnding.text : '';
    template.upgradeText = user.planName == 'Trial' ? 'UPGRADE NOW' : '';
    template.bgColor = trialEnding.bgColor;
    template.templateUrl = getAppUrl_ + '/#/app/templates';
    template.upgradeUrl = getAppUrl_() + '/#/app/account/upgrade';
  }
 
    var page = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Kloudio - My Templates')
      .setWidth(300);
  console.log("after generating page...");
   console.log('Access token = ' + ScriptApp.getOAuthToken());
  return page;
}

function generateReportListPage_()
{
 console.log('generateReportListPage_');
  var template = HtmlService.createTemplateFromFile('ReportList');
  if(checkForExpiration_()){
    showUpgradePage_();
    return null;
  }
  else
  {
    var user = currentUser_();
    var trialEnding = getTrialEndingText_();
    template.trialText = user.planName == 'Trial' ? trialEnding.text : '';
    template.upgradeText = user.planName == 'Trial' ? 'UPGRADE NOW' : '';
    template.bgColor = trialEnding.bgColor;
  }
  

    var page = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Kloudio - Current Reports')
      .setWidth(300);
  console.log("after generating page...");
  return page;
}

function generateUpgradePage_()
{
 console.log('generateUpgradePage_');
  
  var user = currentUser_();
  if(user)
  {
  var template = HtmlService.createTemplateFromFile('UpgradePage');
  template.upgradeText = user.planName == 'Trial' ? 'YOUR TRIAL HAS ENDED' : 'YOUR PLAN IS CANCELED'; 
  template.upgradeButtonText = user.planName == 'Trial' ? 'UPGRADE NOW' : 'RENEW NOW';
  template.planName = user.planName;
  template.installMode = user.installMode;
  template.trial = user.planName == 'Trial';
  template.admin = user.admins && user.admins.length > 0;
  template.contactText = (user.planName == 'Enterprise' || user.planName == 'Pro' || user.planName =='Premium') ? 'Please contact your Kloudio Admin' : '';
  template.upgradeUrl = user.admins && user.admins.length > 0 ? getAppUrl_() +   '#/app/admin/licenses' : '#/app/account/upgrade';
    var page = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Kloudio - Upgrade Now')
      .setWidth(300);
  console.log("after generating upgrade page...");
  return page;
  }
  else
    return null;
}

function showReportList()
{
  var page = generateReportListPage_();
  if(page)
    SpreadsheetApp.getUi().showSidebar(page);
}


function showLoginPage() {
  //console.log("showLoginPage..."  ); 
  SpreadsheetApp.getUi().showSidebar(generateLoginPage_());
}

function showReportsPage() {
  console.log("showReportsPage..."  ); 
 // console.log("Token = " + getToken_());
  if (getToken_() === null) {
    SpreadsheetApp.getUi().showSidebar(generateLoginPage_());
  } else {
        var page = generateQueriesPage_();
      if(page){
          console.log('found page...');
            SpreadsheetApp.getUi().showSidebar(page);       
    }
  }
}

function showUpgradePage_() {
  console.log("showUpgradePage_..."  ); 
  var page = generateUpgradePage_();
  if(page)
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(generateUpgradePage_());

}


function showTemplatePanel() {
  console.log("showTemplatePanel...");
  if(getToken_() === null)
  {
    console.log("generating login page...");
    SpreadsheetApp.getUi().showSidebar(generateLoginPage_());
  }
  else
  {
    var user= currentUser_();
    if(user && user.planName === 'Starter')
    {
      SpreadsheetApp.getUi().alert('Templates are available only in Pro, Premium and Enterprise plans. Please consider upgrading your plan.');
      return;
    }
    console.log("generating templates page...");
    var page = generateTemplatesPage_();
    if(page)
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(page);
  }
}

function refreshAllReports()
{
  console.log('refreshAllReports...');
  
   var reports = getAllReports_();
  
   console.log(reports);
   for(var i in reports)
   {
     console.log("Refreshing report " + reports[i].report);
     var query = getQueryById_(reports[i].report);
     if(query)
       getData(query, reports[i].params,true,reports[i].sheetId,true);
   }
  
   SpreadsheetApp.getActiveSpreadsheet().toast("Refresh all completed","Kloudio");
}


function refreshFormulas() {
  console.log('refreshFormulas...');
  
  SpreadsheetApp.getActiveSheet().getRange(1,1).setFormula(SpreadsheetApp.getActiveSheet().getRange(1,1).getFormula());
  SpreadsheetApp.flush();
}

function openParamsDialog() {
  console.log("openParamsDialog...");
  
  if(checkForExpiration_()){
    showUpgradePage_();
    return;
  }
  
  var query = getCurrentQuery_();
  if(query == null)
  {
     SpreadsheetApp.getUi().alert('The report in the current sheet does not exist in Kloudio. Please make sure the report exists in kloudio and you have access to it. ');
    return;
  }
  if(query.queryParams == null || query.queryParams.length == 0)
  {
    console.log('no params..refreshing data right away...');
    getData(query, [], true);
  }
  else
  {
    var html = HtmlService.createTemplateFromFile('ParameterDialog').evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showModalDialog(html, 'Report Filters');
  }
}

function openScheduleDialog() {
  //console.log("openScheduleDialog...");
  

  var user = currentUser_();
  if(user === null || user == undefined)
    user = getCurrentUser_();
    
  
   if(checkForExpiration_()){
    showUpgradePage_();
    return null;
  }
  
  if(user.planName === 'Starter')
  {
    SpreadsheetApp.getUi().alert("Scheduled Refreshes are only available in Pro, Premium and Enterprise plans. Please consider upgrading your plan.");
    return;
  }

  var query = getCurrentQuery_();
  var template = getCurrentTemplate_();
  if(query === null && template === null)
  {
    console.log('cannot find query...');
     SpreadsheetApp.getUi().alert('Current Sheet does not contain any Kloudio report or template. Please make sure the report/template exists in kloudio and you have access to it. ');
    return;
  }
  console.log('query found...');
  if(query !== null || template !== null)
  {
    var html = HtmlService.createTemplateFromFile('JobDialog').evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .showModalDialog(html, 'Update Schedule - ' + ( query ? query.queryName : template.templateName) );
  }
}


function handleUnAuthorized_()
{
  console.log('handling unauthorization...');
  SpreadsheetApp.getUi().alert('Your Kloudio session has expired!. You will need to login again.');
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
  
  addMenuEntries_();
  showLoginPage();
}

function openColumnsDialog_() {
  console.log("openColumnsDialog_...");
  var html = HtmlService.createTemplateFromFile('ColumnsDialog').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Add Columns');
}

function getSelectValues(query_id, paramName)
{
  console.log("inside getSelectValues_..for:" + paramName);
  
  var options = getParseOptions_();
  console.log(options);
  var token = getToken_();
  
  if(token === null)
  {
    console.log("token is null");
    handleUnAuthorized_();
    return [];
  }
  else
  {
   console.log('going to fetch select values...'); 
    //var url = apiUrl + '/queries?src=gsheets';
    var url = getApiUrl_() + '/select/' + query_id + '/' + paramName;
    console.log("url = " + url);
    try
    {
      var result = UrlFetchApp.fetch(url,options);
      var respCode = result.getResponseCode();
      if(!handleResponse_(respCode))
      {
        return [];
      }
      else
      {
        console.log("After the urlfetch...");
        // console.log(result);
        var o  = Utilities.jsonParse(result.getContentText());
        if(o.success){
          console.log('got select values...');
          console.log(o.data);
          return o.data;
        }
        else
          return [];
      }
    }
    catch(ex)
    {
      console.log(ex);
      return [];
    }
                                                                                                
  }
}


function currentUser_()
{
 var userProperties = PropertiesService.getUserProperties();
 var str =  userProperties.getProperty('CURRENT_USER');
  if(str)
  {
    return JSON.parse(str);
  }
  else
    return null;
}

function handleResponse_(respCode)
{
    
  if(respCode == 401 || respCode == 403 || respCode == 402)
  {
    switch(respCode)
    {
      case 401:
        handleUnAuthorized_();
        break;
      case 402:
        SpreadsheetApp.getUi().alert('Report fetches more than the supported number of rows (50,000). Please try to use filter to reduce the data volume.');
        break;
      case 403:
        showUpgradePage_();
        break;
      default:
        break;
    }
    return false;
  }
  return true;
}

function getCurrentUser_()
{
    console.log("inside getCurrentUser_...");
  //showAlert("getCurrentUser_");

  var options = getParseOptions_();
  //showAlert(options.headers.Authorization);
  var result = UrlFetchApp.fetch(getApiUrl_() + '/me', options);
  var respCode = result.getResponseCode();
  console.log(result);
  
  if(!handleResponse_(respCode))
  {
    return null;
  }
  else
  {
    // showAlert("After the urlfetch..");
    console.log("After the urlfetch...");
    var o  = Utilities.jsonParse(result.getContentText());
    if(o.success)
    {
      console.log('o.success...');
      // showAlert("success...");
      console.log(o.user);
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('CURRENT_USER', JSON.stringify(o.user));
      //showAlert("After saving current user..");
      return o.user;
    }
    else
    {
      return null;
    }
  }
}

function getQueries(queryName)
{
  console.log("inside getQueries_..");
  var options = getParseOptions_();
  console.log(options);
  var token = getToken_();
  console.log(token);
  
  if(token === null)
  {
    console.log("token is null");
    handleUnAuthorized_();
    return {success: false, queries: [], timezones: []};
  }
  else
  {
   console.log('going to fetch queries...'); 
    //var url = apiUrl + '/queries?src=gsheets';
    var url = getApiUrl_() + '/queries?src=gsheets';
    console.log("url = " + url);
    var currentUser = getCurrentUser_();
    try
    {
      var result = UrlFetchApp.fetch(url,options);
      var respCode = result.getResponseCode();
      console.log("responseCode  = " + respCode);
      
      if(!handleResponse_(respCode))
      {
        return {success: false, queries: [], timezones: [], user: currentUser};
      }
      else
      {        
        console.log("After the urlfetch...");
        //console.log(result);
        
        var o  = Utilities.jsonParse(result.getContentText());
        var timezones = getTimeZoneList_();
        if(o.success)
          return {success: true,queries: o.queries, timezones: timezones, user: currentUser};
        else
          return {success: false,queries: [], timezones: [], user: currentUser};
      }
    }
    catch(ex)
    {
      console.log(ex);
      return {success: false, queries: [], timezones: [], user: currentUser};
    }
                                                                                                
  }

}

function getTemplates()
{
  console.log("inside getTemplates_..");
  var options = getParseOptions_();
  console.log(options);
  var token = getToken_();
  console.log(token);
  
  if(token === null)
  {
    console.log("token is null");
    handleUnAuthorized_();
    return {success: false, templates: [], timezones: []};
  }
  else
  {
    var currentUser = getCurrentUser_();
   console.log('going to fetch templates...'); 
    //var result = UrlFetchApp.fetch(apiUrl + '/templates?src=gsheets' , options);
    var result = UrlFetchApp.fetch(getApiUrl_() + '/templates?src=gsheets' , options);
    var respCode = result.getResponseCode();
    
    if(!handleResponse_(respCode))
    {
      return {success: false, templates: [], timezones: [], user: currentUser};
    }
    else
    {
      console.log("After the urlfetch...");
      // console.log(result);
      var o  = Utilities.jsonParse(result.getContentText());
      if(o.success) {
        var timezones = getTimeZoneList_();
        return {success: true, templates: o.templates, timezones: timezones, user: currentUser};
      }
      else
        return {success: false, templates: [], timezones: [], user: currentUser };
    }
                                                                                                
  }

}

function getQueryById_(queryId)
{
  console.log("inside getQueryById_ for " + queryId);
  var options = getParseOptions_();
 
  var result = UrlFetchApp.fetch(getApiUrl_() + '/queries/' + queryId , options);
  var respCode = result.getResponseCode();
  console.log(result);
  
   if(!handleResponse_(respCode))
    {
     return null;
    }
    else
    {
      console.log("After the urlfetch...");
      var o  = Utilities.jsonParse(result.getContentText());
      if(o.success)
      {
        console.log('o.success...');
        console.log(o.query);
        return o.query;
      }
      else
      {
        //SpreadsheetApp.getUi().alert("The query for this table could not be found in kloudio. Please check if the query exists in Kloudio.");
        return -1;
      }
    }

}

function getTemplateById_(tmplId)
{
  console.log("inside getTemplateById_ for " + tmplId);
  var options = getParseOptions_();
  //console.log("payload = " + options.payload);
  //var result = UrlFetchApp.fetch(apiUrl + '/templates/' + tmplId , options);
  var result = UrlFetchApp.fetch(getApiUrl_() + '/templates/' + tmplId , options);
  console.log("After the urlfetch...");
  var respCode = result.getResponseCode();
  if(!handleResponse_(respCode))
  {
    return null;    
  }
  else
  {
    var o  = Utilities.jsonParse(result.getContentText());
    if(o.success)
      return o.template;
    else
    {
      SpreadsheetApp.getUi().alert("The template for this table could not be found in kloudio. Please check if the template exists in Kloudio or you have access to it.");
      return null;
    }
  }
}

function getQueryByName_(queryName)
{
  console.log("inside getQueryByName_ for " + queryName);
  var options = getParseOptions_();
  options.payload = JSON.stringify({"queryName": queryName});
  //console.log("payload = " + options.payload);
  //var result = UrlFetchApp.fetch(apiUrl + '/functions/queryByName' , options);
  var result = UrlFetchApp.fetch(getApiUrl_() + '/functions/queryByName' , options);
  
  var respCode = result.getResponseCode();
  if(!handleResponse_(respCode))
  {
  console.log("After the urlfetch...");
 // console.log(result);
   var o  = Utilities.jsonParse(result.getContentText());
  return o.result;
  }
  else
    return null;
}


function getFrequencies_(){
 var email = Session.getEffectiveUser().getEmail();
 return ['Minute','Hourly','Daily', 'Weekly', 'Monthly'];
 
}



function getScheduleByCurrentSheet(){
  console.log('getScheduleByCurrentSheet');
  var currentSpreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var currentSheetId =  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
console.log("inside getScheduleByQuery for sheet:" + currentSheetId);
  var options = getParseOptions_();
  var result = UrlFetchApp.fetch(getApiUrl_() + '/jobs/' + currentSpreadSheetId + '/' + currentSheetId , options);
  
   var respCode = result.getResponseCode();
  console.log('respCode', respCode);
  if(handleResponse_(respCode))
  {
    console.log("After the urlfetch...");
    var user = getCurrentUser_();
    console.log('user', user);
    var o  = Utilities.jsonParse(result.getContentText());
    console.log(o)
    var tzs = getTimeZoneList_();
    console.log(tzs);
    if(o.success)
      return {schedule: o.job, timezones: tzs, frequencies: getFrequencies_(), user: user};
    else
      return {schedule: null, timezones: tzs, frequencies: getFrequencies_(), user: user};
  }
  else
    return {schedule: null, timezones: tzs, frequencies: getFrequencies_(), user: user};
    
}

function saveAuditJobForCurrentSheet_(jobRec) {
  
  try {
    console.log('saveAuditJobForCurrentSheet_...');
    var currentSheetId = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
    var auditJobStr = JSON.stringify(jobRec);
    setDocumentProperty_('AUDIT_' + currentSheetId, auditJobStr);
  }
  catch(ex){
    console.log(ex);
  }
}

function deleteAuditJobForSheet_(sheetId){
  
  try{
    console.log('deleteAuditJobForSheet_');
    deleteDocumentProperty_('AUDIT_' + sheetId);
  }
  catch(ex){
    console.log(ex);
  }
}

function getAuditJobForCurrentSheet_(currentSheetId) {
  
  try{
    console.log('getAuditJobForCurrentSheet_...');
    var jobStr = getDocumentProperty_('AUDIT_' + currentSheetId);
    var job = (jobStr == null || jobStr == '') ? null : JSON.parse(jobStr);
    return job;
  }
  catch(ex){
    console.log(ex);
    return null;
  }
}

function getParamsForCurrentSheet_(pSheetId) {
  try{
    console.log('getParamsForCurrentSheet_...');
    var currentSheetId =  pSheetId ? pSheetId : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
    var currentSSId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var paramStr = getDocumentProperty_('PARAMS_' + currentSheetId);
    if(paramStr == null || paramStr == '')
    {
      return [];
    }
    else
    {
      return JSON.parse(paramStr);
    }
    
  }
  catch(ex)
  {
    console.log(ex);
    return null;
  }
  
}

function getTimeZones_(){
  console.log('timezones...');
  var tzs = moment.tz.names();
  //console.log(tzs);
  return tzs;
}


function updateSchedule(schedule) {
  try{
  console.log('updateSchedule...');
   var currentSpreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var currentSheetId =  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
  console.log("inside updateSchedule_ for sheet:" + currentSheetId);
  console.log(schedule);
  var query = getCurrentQuery_();
  console.log('after getting current query...' + query.queryName);
  var queryId = query.key;
  console.log('query key = ' + queryId);
 var options = getParseOptions_();
  console.log('debug1');
  options['method'] = schedule.job_id === undefined ? "post" : "put";
   console.log('debug2');
        
  var report_params = getParamsForCurrentSheet_();
  console.log('Report Params = ' + report_params);
    
    options.payload = JSON.stringify({"job":{"job_type": schedule.job_type ? schedule.job_type: 'REPORT', "frequency": schedule.frequency, "sheet_id": currentSheetId.toString(), "spreadsheet_id": currentSpreadSheetId,"report_id": queryId,
                                           "report_name":query.queryName,  "day": schedule.day, "hour": schedule.hour,"min": schedule.minute, "am_pm": schedule.am_pm,
                                           "timezone": schedule.timezone, "enabled": schedule.enabled, "report_params": report_params, "email_on_success": schedule.email_on_success}});
   console.log('debug3');
  console.log(options.payload);
  var url =  getApiUrl_() + '/jobs';
  var result = UrlFetchApp.fetch(url , options);
  //console.log('Result is ');
    // console.log(result);
    var respCode = result.getResponseCode();
    if(!handleResponse_(respCode))
    {
      SpreadsheetApp.getUI().alert("An error occured while updating the schedule information. Please try again later");
      return;
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("Schedule update succesful","Kloudio");
  }
  catch(ex){
    SpreadsheetApp.getUI().alert("An error occured while updating the schedule information. Please try again later");
    console.log(ex.message);
  }
  
}

// Helper function that puts external JS / CSS into the HTML file.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


function getCurrentQuery_()
{
  try
  {
    console.log("getCurrentQuery_...");

    //var tblName = getTableForCurrentCell_();
    var nr = getNamedRangeForCurrentSheet_();
    if(nr == null)
    {
      console.log('No named range found...');
      return null;
    }
    var tblName = nr.getName();
    if(tblName !== null)
    {
      console.log("tblName = " + tblName);
      return getQueryForTable_(tblName);
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}

function getCurrentTemplate_(){
  
  try {
    var namedRange = getNamedRangeForCurrentSheet_();
    if(namedRange !== null)
    {
      var tableName = namedRange.getName();
      return getTmplForTable_(tableName);
    }
  }catch(ex) {
    console.log(ex);
  }
  return null;
}

function getCurrentQueryWithParams()
{
  try{
    console.log('getCurrentQueryWithParams_...');
    var currentQuery = getCurrentQuery_();
    var pVals = getParamsForCurrentSheet_();
    console.log('Report params = ' + pVals);
    var selectData = {};
    
    if(currentQuery.queryParams && currentQuery.queryParams.length > 0)
    {
      currentQuery.queryParams.forEach(function(param){
        if(param.paramType == 'Select' && param.paramValue) {
          if(startsWith_(param.paramValue.toUpperCase(),'SELECT '))
          {
            console.log( param.paramName + '...this is a select query...');
            selectData[param.paramName] = getSelectValues(currentQuery.key,param.paramName);
          }
          else
          {
            console.log('comma separated paramvalue = ' + param.paramValue);
            selectData[param.paramName] = param.paramValue !== '' ? param.paramValue.split(',').map(function(item){return {display: item, value:item};}) : [];
          }
        }
      });
    }
    return {query: currentQuery, paramVals: pVals, selectData: selectData};
  }
  catch(ex)
  {
    console.log(ex);
    return {query: null, paramVals: null};
  }
}

function getNotDisplayedCols_(query)
{
  try
  {
    console.log("getNotDisplayedCols_...");
    //var nRange = getNamedRange_ForCurrentCell_();
    var nRange = getNamedRangeForCurrentSheet_();
    var notDisplayedCols = [];
    if(nRange != null)
    {
      console.log("nRange not null..");
      var hdrCols = nRange.getRange().getSheet().getRange(nRange.getRange().getRow(),nRange.getRange().getColumn(),1,nRange.getRange().getNumColumns()).getValues()[0];
      query.availableCols.filter(function(el) { return el.visible;})
        .forEach(function(col)
                 {
                   if(!_.contains(hdrCols,col.columnName.toLowerCase()))
                     notDisplayedCols.push({columnName: col.columnName, caption: col.caption, visible: false});
                 }
                );
    }
    else
      console.log("nRange is null..");

     return notDisplayedCols;                                                                      
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return null;
}

function getQueryAndCols_()
{
  try
  {
    console.log("getQueryAndCols_...");
    var query = getCurrentQuery_();
    console.log("Got query = " + query.queryName);
    var cols = getNotDisplayedCols_(query);
    console.log("Got cols, count = " + cols.length);
    return {query: query, cols: cols};
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}

function addColsToTable_(query, colsToBeAdded)
{
  try
  {
    console.log("addColsToTable_...");
    if(colsToBeAdded != null && colsToBeAdded.length > 0)
    {
      console.log("there are cols to be added..");
      //var nmRange = getNamedRange_ForCurrentCell_();
      var nmRange = getNamedRangeForCurrentSheet_();
      colsToBeAdded.forEach(function(col)
                            {
                              console.log("for Col: " + col.caption);
                              var nRange = nmRange.getRange();
                              nRange.getSheet().insertColumnAfter(nRange.getColumn());
                              nRange = nmRange.getRange();
                              nRange.getCell(1,2).setValue(col.columnName);
                            }
                            );
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}


function getTableForCurrentCell_()
{
  try
  {
    console.log("getTableForCurrentCell_...");
    var namedRange = getNamedRange_ForCurrentCell_();
    if(namedRange != null)
      return namedRange.getName();
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return null;
}

function getNamedRange_ByName_(nrName)
{
  try
  {
    console.log("getNamedRange_ByName_...");
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   
    var nra = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getNamedRanges();
   
    if(nra != null && nra.length > 0)
    {
      var namedRange =  _.find(nra, function(nr)
              {
               return nr.getName() == nrName;
              }
              );
      
      return namedRange;
    }

  }
  catch(ex)
  {
    console.log(ex.message);
  }
}

function getNamedRangeByName_(nrName)
{
  try
  {
    console.log("getNamedRangeByName_ for: " + nrName);
    var nra = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
    
    if(nra !==null && nra.length > 0)
    {
      for(var i in nra)
      {
        var nr = nra[i];
        if(nr.getName() === nrName){
          console.log('Found namedRange ..' + nrName);
          return nr;
        }
      }
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}


function getNamedRangeForCurrentSheet_(sheet)
{
  try
  {
    console.log("getNamedRangeForCurrentSheet_...");
    var ss = sheet ? sheet : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    console.log("Active spreadsheet = " + ss.getName());
    var cell = ss.getActiveCell();
    var nra = ss.getNamedRanges();
    
    if(nra !== null && nra.length > 0)
    {
      for(var i in nra) {
        if(startsWith_(nra[i].getName(), 'DO_NOT_DELETE_KLOUDIO'))
        {
          return nra[i];
        }
      }
      return nra[0];
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return null;
}

function getNamedRange_ForCurrentCell_()
{
  try
  {
    console.log("getNamedRange_ForCurrentCell_...");
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    console.log("Active spreadsheet = " + ss.getName());
    var cell = ss.getActiveCell();
    var nra = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getNamedRanges();
    var r = cell.getRow();
    var c = cell.getColumn();
    console.log("r = " + r + " c = " + c);
    if(nra != null && nra.length > 0)
    {
      console.log("There are named ranges..");   
      var namedRange =  _.find(nra, function(nr)
              {
                var nrRange = nr.getRange();
                var nrr =  nrRange.getRow();
                var nrc = nrRange.getColumn();
                return nrr <= cell.getRow() && nrc <= cell.getColumn() && (nrr + nrRange.getNumRows()) > cell.getRow() && (nrc + nrRange.getNumColumns()) > cell.getColumn(); 
              }
              );
      
      return namedRange;
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}

function selectSheet(sheetName)
{
  console.log('selectSheet');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.activate();
}

function getJobsInfo_()
{
  console.log('getJobsInfo_...');
   var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  try{
    var url = getApiUrl_() + '/sheets/' + spreadsheetId;
    var options = getParseOptions_();
      var result = UrlFetchApp.fetch(url,options);
      var respCode = result.getResponseCode();
      console.log("responseCode  = " + respCode);
      
      if(!handleResponse_(respCode))
      {
        return null;
      }
      else
      {
        console.log("After the urlfetch...");
        //console.log(result);
        var o  = Utilities.jsonParse(result.getContentText());
        var timezones = getTimeZoneList_();
        if(o.success){          
          return o.sheetsInfo;
        }
      }    
  }
  catch(ex)
  {
    console.log(ex);
  }
  return null;
}

function getReportsInfo_(reportIds)
{
  console.log('getReportsInfo_');
   var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

    var url = getApiUrl_() + '/sheets/reportInfo';
    var options = getParseOptions_();
    options['method'] = 'post';
    options.payload = JSON.stringify({report_ids: reportIds});    
    var result = UrlFetchApp.fetch(url,options);
    var respCode = result.getResponseCode();
    console.log("responseCode  = " + respCode);
    
    if(!handleResponse_(respCode))
    {
      return null;
    }
    else
    {
      console.log("After the urlfetch...");
      //console.log(result);
      var o  = Utilities.jsonParse(result.getContentText());
      var timezones = getTimeZoneList_();
      if(o.success){          
        return o.reports;
      }
    }    
  return null;
}

function getCurrentReports(){
  console.log('getCurrentReports');
  try {
    var currentSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var report_ids = [];
    var reports =[];
    for(var i=0; i < currentSheets.length; i++)
    {
      var nra = currentSheets[i].getNamedRanges();
      for(var j=0; j< nra.length; j++)
      {
        var tblName = nra[j].getName();
        var report_id = getDocumentProperty_(tblName);
        var tblRunInfo = getDocumentProperty_(tblName + '_RUN');
        if(report_id)
        {
          report_ids.push(report_id);
          reports.push({spreadsheetId: spreadsheetId ,sheetId: currentSheets[i].getSheetId(),sheetName: currentSheets[i].getName(),reportId: report_id, runInfo:tblRunInfo ? JSON.parse(tblRunInfo) : tblRunInfo });
        }
      }
    }
    var reportsInfo = getReportsInfo_(report_ids);
    var jobsInfo = getJobsInfo_();
    return {reports: reports, reportsInfo: reportsInfo, jobsInfo: jobsInfo};    
  }
  catch(ex)
  {
    console.log(ex); 
    if(ex.name == 'UnAuthorizedException')
    {
       handleUnAuthorized_();
    }
  }  
  return null;
}

function getDataPvt_(query, parameters,noHeaders)
{
  try
  {
    console.log("getDataPvt_");
    var options = getParseOptions_();
    var selectCols =  query.sqlMode ? [] : query.availableCols.filter(function(el) { return el.visible;}).map(function(col){return col.columnName  ;}).join(",");
    var params = parameters;
    options.payload = JSON.stringify({"query": {"query_params": params, "select_cols": selectCols, "src":"gsheets","token":getToken_() }});
    console.log(options.payload);
    options["method"] = "post";
    //var url = apiUrl + '/queries/' + query.key + '/execute';
    var url = getApiUrl_() + '/queries/' + query.key + '/execute';
    console.log('url = ' + url);
    if(noHeaders != undefined)
    {
      console.log("noHeaders not undefined..");
      return ImportJSON_(url,options,noHeaders);         
   
    }
    else
    {
      console.log("noHeaders undefined..");
      return ImportJSON_(url,options);
    }
  }
  catch(ex)
  {
    console.log(ex.message);
    throw ex;
  }
}


function executeFunction_(fnName, parameters)
{
  try
  {
    console.log("executeFunction_");
    var options = getParseOptions_();
    options.payload = JSON.stringify({"fn": fnName, "params": parameters, "src": "gsheets"});
    console.log(options.payload);
    options["method"] = "post";
    var url = getFnUrl_() + '/functions/' + fnName + '/execute';
    console.log('url = ' + url);
     var result = UrlFetchApp.fetch(url,options);
    console.log(result);
      console.log('got response from server...');
      var object = JSON.parse(result.getContentText());
      console.log('object is...');
      console.log(object);
      return object;
  }
  catch(ex)
  {
    console.log(ex.message);
    return {error: ex.message};
  }
}

function insertAsFormula(query, params) {
  try{
    console.log("insertAsFormula...");
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    activeSpreadsheet.insertSheet();
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var startRow = 1;
    var formulaArray = ['=Kloud("' + query.key + '"'];
    params.forEach(function(p) {
        formulaArray.push('"' + p.param_name + '"');
        
        var qps = query.queryParams && query.queryParams.length > 0 ? query.queryParams.filter(function(ele) {
          return ele.paramName === p.param_name;
        }) : null;
        var qp = qps && qps.length > 0 ? qps[0] : null;
        
        if(qp) {
          activeSheet.getRange(startRow, 1).setValue(qp.paramDisplayName);
          switch(qp.paramType) {
            case 'Number':
              formulaArray.push('B' + startRow);
              activeSheet.getRange(startRow, 2).setValue(p.param_val);
              break;
            case 'Text':
            case 'Date':
              // formulaArray.push('"' + p.param_val + '"');
              formulaArray.push('B' + startRow);
              activeSheet.getRange(startRow, 2).setValue(p.param_val);
              break;
            default:
              // formulaArray.push('"' + p.param_val + '"');
              formulaArray.push('B' + startRow);
              activeSheet.getRange(startRow, 2).setValue(p.param_val);
              break;
          }
           startRow = startRow + 1;
        }
        else {
           // formulaArray.push(p.param_val);
        }
       
      });
     var formulaCell = activeSheet.getRange(startRow, 1);
      var fullFormula = formulaArray.join(',') + ')';
      console.log('FullFormula: ' + fullFormula);
      formulaCell.setFormula(fullFormula);
      return {success: true}; 
  } catch(ex) {
    console.log(ex);
    return {success: false};
  }
}

function getData(query, parameters,refreshData,sheetId,refreshAll, formula)
{
  try
  {
     var params = parameters  ? parameters : [];
    
    console.log("queryparams:", JSON.stringify(params));
    
    if(formula) {
      return insertAsFormula(query, params);
    }
    else {
      console.log("run as report...");
    var response = getDataPvt_(query,params);
    if(!response.success) {
      console.log('Data fetch failed...');
      SpreadsheetApp.getUi().alert(response.message);
     return {success: false};
    }
    
    var data = response.data;
    
    if(refreshData)
    {
      console.log("this is refresh mode...");
      
      if(!!!data[0])
      {
        SpreadsheetApp.getUi().alert('No data found for the filter criteria!');
        return {success: false};
      }
      //console.log(data);
      numCols= data[0].length;
      numRows = data.length;
      console.log('numRows = ' + numRows);
      if(numRows == 0)
      {
        SpreadsheetApp.getUi().alert('No data found for the filter criteria!');
        return {success: false};     
      }
      
      var currentSpreadsheetId = sheetId ? sheetId : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
      refreshResults_(query,data,numCols,currentSpreadsheetId);
      
      
      var paramStr = JSON.stringify(params);
      console.log('ParamStr = ' + paramStr);
      setDocumentProperty_('PARAMS_' + currentSpreadsheetId , paramStr);
      
      console.log('after storing query params...');
      if(!!!refreshAll)
        SpreadsheetApp.getActiveSpreadsheet().toast("Data Refresh completed","Kloudio");
    }
    else
    {
      console.log("this is new mode...");
      if(!!!data[0])
      {
        SpreadsheetApp.getUi().alert('No data found for the query criteria!');
        return {success: false};
      }      
      var numRows = data.length;
      console.log('numrows = ' + numRows);
   
      if(numRows > 1)
      {
        console.log('rows > 1');
        var firstEmptyRow = SpreadsheetApp.getActiveSpreadsheet().getLastRow() + 1;
        if(firstEmptyRow > 1)
        {
          console.log("current sheet is not empty...");
          SpreadsheetApp.getActiveSpreadsheet().insertSheet(generateSheetNameForQuery_(query.queryName));
        }
        
        //console.log(data);
        numCols = data[0].length;
        console.log('numCols = ' + numCols);
        
        renderResults_(query,data, numCols);
        var currentSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
      
        var paramStr = JSON.stringify(params);
        console.log('ParamStr = ' + paramStr);
        setDocumentProperty_('PARAMS_' + currentSpreadsheetId , paramStr);
        console.log('after storing query params...');
        return {success: true}
      }
      else
      {
        SpreadsheetApp.getUi().alert('No rows found for the query criteria!');
        return {success: false};
      }
    }
    
    return {success: false};
    }
  }
  catch(ex)
  {
    console.log(ex);
    SpreadsheetApp.getUi().alert('An error occurred while fetching the data. Please try again later!');
  }
   return {success: false};
}

function updateResults_(query,data,numCols)
{
  try
  {
    console.log("updateResults...");
    
    //1. Remove rows that does not exist in the database from the excel range....
    //2. Add new ids to the end of the range..
    //4. Now for each column, get the array from the query data and populate in chunk in the order of the ids
    //4. 
    //var namedRange = getNamedRange_ForCurrentCell_();
    var namedRange = getNamedRangeForCurrentSheet_();
    var rng = namedRange.getRange();
    var sheet = rng.getSheet();
    var rngColumn = rng.getColumn();
    var dataRowCount = data.length;
    var currentRowCount = rng.getNumRows() - 1;
    
    //var keyColumnCaption = Enumerable.from(query.availableCols).where(function(c) { return c.columnName == query.outputInfo.keyColumn;}).select(function(c) { return c.caption }).toArray()[0];
    
   // console.log("Key Column Caption = " + keyColumnCaption);
    
    var tableName = namedRange.getName();
    
    var keyColumnIndex = findColIndexInTable_(tableName, query.outputInfo.keyColumn.toLowerCase());
    
    console.log("Key Column Index = " + keyColumnIndex);
    
    
    var idCols2D = rng.getSheet().getRange(rng.getRow() + 1 , rng.getColumn() + keyColumnIndex - 1, rng.getNumRows() - 1, 1).getValues();
    var idCols = [];
    idCols2D.forEach(function(item) { idCols.push(item[0]);});
    
    console.log("idCols length = " + idCols.length);
    console.log("idCols = " + idCols);
    
    var enumerable = Enumerable.from(data);
    var idArray =  enumerable.select(function(x){return x[query.outputInfo.keyColumn.toLowerCase()];}).toArray();
    var colNames = Object.keys(enumerable.first());
    console.log("colNames = " + colNames);
    var colNameIndices = [];
    colNames.forEach(function(key) { colNameIndices.push({colName: key, colIndex: findColIndexInTable_(tableName, key)}) });
    
   // console.log("colNameIndices" + colNameIndices);
    
    console.log("idArray = " + idArray);
    var newIds = _.difference(idArray,idCols);
    var idsToBeDeleted = _.difference(idCols,idArray);
    
    if(newIds != null && newIds.length > 0)
    {
      console.log("new Ids exists  and count = " + newIds.length);
      resizeNamedRange_(namedRange, newIds.length + currentRowCount);
      namedRange = getNamedRange_ByName_(namedRange.getName());
      rng = namedRange.getRange();
      var newIdRange = rng.getSheet().getRange(rng.getRow() + 1 , rng.getColumn() + keyColumnIndex - 1, newIds.length, 1);
      var newIdData = [];
      newIds.forEach(function(id) { newIdData.push([id])});
      newIdRange.setValues(newIdData);
      console.log("After setting new Ids...");
    }
    
    rng = namedRange.getRange();
    idCols2D = rng.getSheet().getRange(rng.getRow() + 1 , rng.getColumn() + keyColumnIndex - 1, rng.getNumRows() - 1, 1).getValues();
    idCols = [];
    idCols2D.forEach(function(item) { idCols.push(item[0]);});
    var row = 2;
    console.log("idColsLength = " + idCols.length);
    
     idCols.forEach(function(id)
                   {   
                     console.log("For id = " + id);
                     var dataObject = enumerable.where("$." + query.outputInfo.keyColumn.toLowerCase() + " == '" + id + "'").toArray()[0];
                     console.log(dataObject);
                     if(dataObject != null)
                     {
                       console.log("dataObject not null..");
                         //update the row here...
                       
                       
                       colNameIndices.forEach(function(col)
                                              {
                                                if(col.colIndex != -1 && col.colName.toLowerCase() != query.outputInfo.keyColumn.toLowerCase())
                                                {
                                                  var val = dataObject[col.colName];
                                                  console.log("val of col: " + col.colName + " = " + val);
                                                  rng.getCell(row,col.colIndex).setValue(val); 
                                                }
                                              }
                                              );
                       
                     }
                     else
                     {
                       
                     }
                     row++;
                   }
                   );
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
}

function createReportHeader_(queryName) {
  
  try {
    console.log("Inside createReportHeader_...");
    var ssActive = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastCol= ssActive.getLastColumn();
    var myRange  = ssActive.getRange(1,1, 1, lastCol);
    myRange.setBackground('#2f82c9');
    myRange.setFontWeight('Bold');
    myRange.setFontColor('#fff');
    myRange.setFontSize(12);
    myRange.merge();
    myRange.setValue(queryName);
    console.log("after header formatting...");
    
  }catch(ex) {
    console.log(ex);
  }
}

function renderResults_(query,data, numCols)
{
  try
  {
    console.log("Inside renderResults");
    var ssActive = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ssId = ssActive.getSheetId();
    var numRows = data.length;
    console.log("activeSheet = " + ssActive.getName());
    console.log("numCols = " + numCols);
    console.log("numRows = " + numRows);
  
    //console.log(data);
   
    var myRange  = ssActive.getRange(1,1, numRows, numCols);
    console.log("after getting range");
    myRange.setValues(data); 
    console.log("after setting values...");
    if(!query.sqlMode)
    {
      console.log("query not in sql mode...");
      var hdrRange  = ssActive.getRange(1,1,1,numCols);
      
      var captions = query.selectedCols.map(function(col){return col.caption || col.columnName;});
      console.log('Captions');
      console.log(captions);
      var arr = new Array(1);
      arr[0] = captions;
      hdrRange.setValues(arr);
    }
    var tblName = generateTblNameForQuery_(query.queryName);
    // var tblName = K + ssId;
    console.log("Tablename = " + tblName);
    createNamedRange_(tblName,myRange);
    console.log("After creating named range...");
    // createReportHeader_(query.queryName);
    // console.log("After creating report header...");
    saveTableInfo_(tblName,query);
    console.log("after saving table info...");
    saveTableRunInfo_(tblName,{queryId: query.key,queryName: query.queryName, sheetId: ssActive.getSheetId(), lastRefreshed: new Date(),rowsRefreshed: data ? data.length - 1 : 0});
    
    console.log("After saving run info...");  
  }
  catch(ex)
  {
    console.log(ex);
  }

}

function toDataTable_(query,data, numCols)
{
  try
  {
    console.log("toDataTable...");
    // console.log(data);
    console.log("numRows = " + data.length);
    console.log("numCols= " + numCols);
    var dataTable = {};
    var captions = [];
    var namesToCaptions = {};

    if(!query.sqlMode)
    query.selectedCols.forEach(function(col) { 
      namesToCaptions[col.columnName] = col.caption || col.columnName; 
    });
    
    
    for(i = 0; i< data.length; i++)
    {
      //console.log("i = " + i);
      for (j= 0; j < numCols; j++)
      {
        //console.log("j = " + j);
        if(i ==0)
        {
          if(query.sqlMode) //for sql mode col name = caption
            namesToCaptions[data[0][j]] = data[0][j];
          
          dataTable[namesToCaptions[data[0][j]]] = [];
        }
        else
        {
          dataTable[namesToCaptions[data[0][j]]].push(data[i][j]);
        }
      }
    }
    
    return dataTable;
  }
  catch(ex)
  {
    console.log(ex);
  }
}

function getSheetById_(sheetId)
{
  try {
    console.log('getSheetById...');
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i in sheets)
    {
      if(sheets[i].getSheetId() === sheetId)
        return sheets[i];
    }
  }
  catch(ex){
    console.log(ex);
  }
  return null;
}

function refreshResults_(query,data,numCols,sheetId)
{
  try
  {
    console.log("Inside refreshResults");
//    console.log("numRows = " + data.length - 1);
    
    var startTime = new Date();
    var ssActive = sheetId ? getSheetById_(sheetId) :  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    console.log("For sheet: "  + ssActive.getName());
   
    currentDataTable = toDataTable_(query,data,numCols);
    //console.log(currentDataTable);
    
    var endTime = new Date();
    
    var elapsedTime = endTime - startTime ;
    
    console.log("Creating dataTable took :" + elapsedTime);    
    
    
    
   // var namedRange = getNamedRange_ForCurrentCell_();
    var namedRange = getNamedRangeForCurrentSheet_(ssActive);
    
    resizeNamedRange_(namedRange,data.length - 1);
    
    console.log("after resizing nmRange..");
    
    
    if(namedRange != null)
    {
      console.log("namedRange not null...");
      var nrRange = namedRange.getRange();
      var nmHeadersRange = nrRange.getSheet().getRange(nrRange.getRow(), nrRange.getColumn(),1,nrRange.getNumColumns());
      var nmHeaders = nmHeadersRange.getValues()[0];
      var colPosition = nrRange.getColumn();
      var rowPosition = nrRange.getRow();
      nmHeaders.forEach(
        function(hdrCol)
        {
          console.log("for hdrCol = " + hdrCol);
          
          var hdrColData = currentDataTable[hdrCol];

          if(hdrColData !== undefined)
          {
            console.log("colData Found..");
            var colData = [];
            for(i = 0; i < hdrColData.length; i++)
              colData.push([hdrColData[i]]);
            nrRange.getSheet().getRange(rowPosition + 1, colPosition, data.length - 1, 1).setValues(colData);
          }
          colPosition++;
        }
      );
      saveTableRunInfo_(namedRange.getName(),{queryId: query.key,queryName: query.queryName, sheetId: ssActive.getSheetId(), lastRefreshed: new Date(), rowsRefreshed: data ? data.length - 1 : 0});
    }
   
    console.log("After inserting data...");  
  }
  catch(ex)
  {
    console.log(ex.message);
  }

}

function getNamedRange_(nr)
{
  try
  {
    return SpreadsheetApp.getActiveSpreadsheet().getRangeByName(nr);
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}

function generateTblNameForQuery_(queryName)
{
  try
  {
    // var qn = queryName.replace(/[^a-zA-Z0-9_]/g,'_');
    // qn = qn.replace(/ /g, '_');
    var qn = 'DO_NOT_DELETE_KLOUDIO';
    var name = qn + '1';
    var i = 2;
    while(getNamedRange_(name) != null)
    {
      name = qn + i;
      i++;
    }
    
    return name;    
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}

function generateSheetNameForQuery_(queryName)
{
  try
  {
    var qn = queryName + "1";
    var i = 2;
    while(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(qn) != null)
    {
      qn = queryName + i;
      i++;
    }
    
    return qn;    
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}

function createNamedRange_(nr, range)
{
  try
  {
    console.log("createNamedRange...");
    
    if(getNamedRange_(nr) == null)
    {
      SpreadsheetApp.getActiveSpreadsheet().setNamedRange(nr, range);
      console.log('after creating namedrange..');
      range.getSheet().getRange(range.getRow(),range.getColumn(),1,range.getNumColumns()).setBackground('#efefef');
      console.log('after setting bg...');
      range.getSheet().getRange(range.getRow(),range.getColumn(),1,range.getNumColumns()).setFontWeight('bold');
      console.log('after setting font-weight');
      //range.getSheet().getRange(range.getRow(),range.getColumn(),1,range.getNumColumns()).setFontColor("white");
      //range.getSheet().getRange(range.getRow(),range.getColumn(),1,range.getNumColumns()).setBorder(true,true,true,true,true,true, "#888");
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}


function deleteAndShiftUp() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var totalHeight = sheet.getDataRange().getHeight();
  var totalWidth = sheet.getDataRange().getWidth();
  var toDelete = sheet.getActiveRange();
  var firstRow = toDelete.getRow();
  var firstColumn = toDelete.getColumn();
  var lastRow = Math.min(toDelete.getLastRow(), totalHeight);
  var lastColumn = Math.min(toDelete.getLastColumn(), totalWidth);
  var height = lastRow-firstRow+1;
  var width = lastColumn-firstColumn+1;
  if (height>0 && width>0) {
    if (totalHeight>lastRow) {
      sheet.getRange(lastRow+1, firstColumn, totalHeight-lastRow, width)
           .copyTo(sheet.getRange(firstRow, firstColumn, totalHeight-lastRow, width));
    }
    sheet.getRange(totalHeight-height+1, firstColumn, height, width).clear();
    
  }
}



function resizeNamedRange_( nr, numRows)
{
  try
  {
    console.log("resizeNamedRange..");
    var currentRows = nr.getRange().getNumRows() - 1;
    var delta = numRows - currentRows;
    console.log("Delta = " + delta);
    var nrRange = nr.getRange();
    var rngSheet = nrRange.getSheet();
    
    if(delta > 0)
    {
      console.log('delta > 0');
      //var nra = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getNamedRanges()[0];
       rngSheet.setActiveSelection(rngSheet.getRange(nrRange.getRow() + nrRange.getNumRows() - 1, nrRange.getColumn(),1,1));
       rngSheet.insertRowsBefore(nrRange.getRow() + nrRange.getNumRows() - 1, delta);
       var newDataRng = rngSheet.getRange(nrRange.getRow() + 1, nrRange.getColumn(), numRows, nrRange.getNumColumns());
       var newRng = nrRange.getSheet().getRange(nrRange.getRow(), nrRange.getColumn(), numRows + 1, nrRange.getNumColumns());
       clearValues_(newDataRng);
       nr.setRange(newRng);      
    }
    else
    {
      
        if(currentRows > 1)
        {
          console.log("currentRows > 1");
          //var rngToDelete = rngSheet.getRange(nrRange.getRow() + nrRange.getNumRows() + delta, nrRange.getColumn(), 0-delta, nrRange.getNumColumns());
          if(delta < 0)
            rngSheet.deleteRows(nrRange.getRow() + nrRange.getNumRows() + delta, 0-delta);         
          //deleteAndShiftUp();
          var newDataRng = rngSheet.getRange(nrRange.getRow() + 1, nrRange.getColumn(), numRows, nrRange.getNumColumns());
          var newRng = rngSheet.getRange(nrRange.getRow(), nrRange.getColumn(), numRows + 1, nrRange.getNumColumns());
          clearValues_(newDataRng);
          nr.setRange(newRng);           
        }
        else
        {
          console.log("currentRows <= 1");
          clearValues_(rngSheet.getRange(nrRange.getRow() + 1, nrRange.getColumn(), 1, nrRange.getNumColumns()));
        }
      }    
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}

function clearValues_(range) {
  var formulas = range.getFormulas();
  range.clearContent();
  range.setFormulas(formulas);
}

function clearNamedRange_(namedRange, startRow)
{
  try
  {
    console.log("clearNamedRange...");
    if(namedRange != null)
    {
      if(startRow != undefined && startRow != null)
      {
        var r = namedRange.getRow();
        console.log("Nr row = " + r );
        var newRange = namedRange.getSheet().getRange(r + startRow, namedRange.getColumn(),namedRange.getNumRows() - startRow, namedRange.getNumColumns());
        clearValues_(newRange);
        console.log("After clearing nr");
      }
      else
      {
        namedRange.clear();
      }
    }
  }
  catch(ex)
  {
    console.log(ex);
  }
}

function getDocumentProperty_(propertyName)
{
  try
  {
    console.log("getDocumentProperty for: " + propertyName);
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty(propertyName);
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  return null;
}

function setDocumentProperty_(propertyName, propertyValue)
{
  try
  {
    console.log("setDocumentProperty for: " + propertyName + ' and value = ' + propertyValue);
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty(propertyName, propertyValue);
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}

function deleteDocumentProperty_(propertyName)
{
  try
  {
    console.log("deleteDocumentProperty for: " + propertyName);
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.deleteProperty(propertyName);
  }
  catch(ex)
  {
    console.log(ex.message);
  }
}


function getTableRunInfo_(tableName)
{
  try
  {
    console.log('getTableRunInfo for ' + tableName);
    var dataStr =  getDocumentProperty_(tableName + '_RUN');
    if(dataStr)
      return JSON.parse(dataStr);
  }
  catch(ex)
  {
    console.log(ex);
  }
  
  return null;
}

function saveTableRunInfo_(tableName, data)
{
  try
  {
    console.log("saveTableRunInfo for " + tableName);
    setDocumentProperty_(tableName + '_RUN', JSON.stringify(data));
  }
  catch(ex)
  {
    console.log(ex);
  }
}

function getAllReports_()
{
  try
  {
    console.log("getAllReports_..");
    var documentProperties = PropertiesService.getDocumentProperties();
    var propKeys = documentProperties.getKeys();
    var reports = [];
    for(var i in propKeys)
    {
      var key = propKeys[i];
      if(startsWith_(key, 'TMPL_') || startsWith_(key, 'PARAMS_') || endsWith_(key, '_RUN'))
        continue;
      var nr = getNamedRangeByName_(key);
      if(nr !== null)
      {
         console.log('nr is not null...');
         var sheetId = nr.getRange().getSheet().getSheetId();
        console.log('sheetId : ' + sheetId);
         var params  = getParamsForCurrentSheet_(sheetId);
        console.log('params');
        console.log(params);
        reports.push({report: documentProperties.getProperty(key), params: params, sheetId: sheetId});                 
      }      
    }
    return reports;
  }
  catch(ex)
  {
    console.log(ex);
  }
  return null; 
}

function saveTableInfo_(tableName, query)
{
  try
  {
    console.log("saveTableInfo for " + tableName);
    setDocumentProperty_(tableName, query.key);
  }
  catch(ex)
  {
    console.log(ex);
  }
}

function saveTableInfoForTmpl_(tableName, tmpl)
{
  try
  {
    console.log("saveTableInfoForTmpl for " + tableName);
    setDocumentProperty_('TMPL_' + tableName, tmpl.key);
  }
  catch(ex)
  {
    console.log(ex);
  }
}


function getTmplIdForTbl_(tableName)
{
  try
  {
    console.log("getTmplIdForTbl...");
    
    var tmplId = getDocumentProperty_('TMPL_' + tableName);
    if(tmplId != null)
    {
      console.log("found tblInfo..");
      return tmplId;       
    }
    else
    {
      console.log("tblInfo not found..");
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return null;
}

function getTmplForTable_(tableName)
{
  try
  {
    console.log("getTmplForTable...");
    
    var tmplId = getDocumentProperty_('TMPL_' + tableName);
    if(tmplId != null)
    {
      console.log("found tblInfo..");
      return getTemplateById_(tmplId);       
    }
    else
    {
      console.log("tblInfo not found..");
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return null;
}

function getQueryForTable_(tableName)
{
  try
  {
    console.log("getQueryForTable...");
    
    var queryId = getDocumentProperty_(tableName);
    if(queryId != null)
    {
      console.log("found tblInfo..");
      var retVal =  getQueryById_(queryId);  
      if(retVal === -1 )
        return null;
      else
        return retVal;
    }
    else
    {
      console.log("tblInfo not found..");
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return null;
}

function getQueryCol_(query,colName)
{
  try
  {
    console.log("getQueryCol..");
   var dbCol =  _.find(query.availableCols, function(col)
           {
             return col.columnName == colName;
           }
           );
    
    return dbCol;
  }
  catch(ex)
  {
    console.log(ex);
  }
  
  return null;
}

function findColIndexInTable_(tableName, columnName)
{
  try
  {
    console.log("findColIndexInTable...");
    var namedRange = getNamedRange_(tableName);
    if(namedRange != null)
    {
      console.log("got namedRange...");
      var query = getQueryForTable_(tableName);
      if(query == null)
      {
        console.log("Cannot get query for the bound table: " + tableName);
        return;
      }
      
      var hdrRange  = namedRange.getSheet().getRange(namedRange.getRow(), 1, namedRange.getColumn(),namedRange.getNumColumns());
      if(hdrRange != null)
      {
        console.log("Got hdrRange...");
        for(i = 1; i<= hdrRange.getNumColumns(); i++)
        {
          var colValue = hdrRange.getCell(1,i).getValue();
          //console.log("For col: " + colValue);
          if(colValue == columnName)
          {
            return i;
          }
          
        }
      }
    }
  }
  catch(ex)
  {
    console.log(ex.message);
  }
  
  return -1;
}

function getFirstEmptyRowWholeRow_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var row = 0;
  for (var row=0; row<values.length; row++) {
    if (!values[row].join("")) break;
  }
  return (row+1);
}

function getTemplateData_(tableName,selection)
{
  try
  {
    console.log('getTemplateData');
    var tmpl = getTmplForTable_(tableName);
    
    if(tmpl == null)
    {
      console.log("cannot find template for table: " + tableName);
      SpreadsheetApp.getUi().alert("This sheet does not contain any template. Please select a sheet containing the template to upload.");
      return null;
    }
    
    CurrentTemplate = tmpl;
    
    var namedRange = getNamedRangeForCurrentSheet_();
    
    if(namedRange == null)
    {
      console.log('cannot find namedRange in current sheet');
      SpreadsheetApp.getUi().alert("This sheet does not contain any template. Please select a sheet containing the template to upload.");
      return null;
    }
    
    var rng = namedRange.getRange();
    
    if(rng == null)
    {
      console.log("cannot get range for namedRange: " + tableName);
      SpreadsheetApp.getUi().alert("This sheet does not contain any template. Please select a sheet containing the template to upload.");
      return null;
    }
    
    var numCols = rng.getNumColumns();
    var numRows = rng.getNumRows();
    console.log("number of rows in rng = " + numRows);
    console.log("number of cols in rng = " + numCols);
    
    var cellsData = rng.getValues();
    var colCaptions = cellsData[0];
    var colNames = colCaptions.map( 
      function(item) { 
        if(item == 'Status' || item == 'Message') return item;
        var col = _.find(tmpl.selectedCols, function(col){
          return col.caption == item;
      }); 
        return col ? col.columnName : '';
      }
      );
                                                 
    var lastRow = getFirstEmptyRowWholeRow_() - 1;
    console.log('range lastRow = ' + lastRow);
    
    var rows = [];
    if(selection)
    {
      rng = SpreadsheetApp.getActiveRange();
      rows = rng.getSheet().getRange(rng.getRow(), 1, rng.getNumRows(), numCols).getValues();
    }
    else
    {
      rows = rng.getSheet().getRange(2, 1, lastRow - 1, numCols).getValues();
    }
    
    
   /* var data = [];
    for (var r = 0, l = rows.length; r < l; r++) {
      var row = rows[r];
      var record = {};
      for (var p in colNames) {
        if(p >1 && colNames[p] !== '')
          record[colNames[p]] = row[p];
      }
      data.push(record);
    }
    return data;
    */
     return {rows: rows, cols: colNames};
    
  }
  catch(ex)
  {
    console.log(ex);
  }
  return [];
}

function validateData_(tmpl, data,rng,startRow)
{
  try
  {
    console.log('validateData');
    var valid  = true;
    
    var reqCols = tmpl.selectedCols.filter(function(col){ return col.required; });
    var expressionValidations =  tmpl.validations.filter(function(item){return item.validationType == 'Expression' && item.enabled;});
    if(reqCols.length > 0 || expressionValidations.length > 0)
    {
      console.log('validating required columns');
      var i = startRow;
      data.forEach(function(row){
        console.log( row);
        var errors = [];
        reqCols.forEach(function(col){
          console.log('for col: ' + col.columnName);
          var val = row[col.columnName];
          console.log('value = ' + val);
          if(val == undefined || val == null || val == '')
          {
            errors.push(col.columnName + ' is required');
          }
        });
        //validate the expression validations
        expressionValidations.forEach(function(validation) {
          
          var stmt = '';
          var evaluation = false;
          switch(validation.operator)
          {
            case '>':
              evaluation = row[validation.column] > validation.compareValue;
              break;
            case '>=':
              evaluation = row[validation.column] >= validation.compareValue;
              break;              
            case '=':
              evaluation = row[validation.column] == validation.compareValue;
              break;              
            case '<':
              evaluation = row[validation.column] < validation.compareValue;
              break;              
            case '<=':
              evaluation = row[validation.column] <= validation.compareValue;
              break;              
            case 'Contained In':
              var valsArray = validation.compareValue.split(',');
              evaluation = valsArray.indexOf(row[validation.column]) > -1;
              console.log('evaluation =' );
              console.log(evaluation);
              break;
            default:
              break;
          }
          if(!evaluation)
          {
            errors.push((validation.message == undefined || validation.message == '') ? 'Validation Failed' : validation.message);
          }
            
        });
        
        if(errors.length > 0)
        {
          rng.getCell(i,1).setValue('invalid');
          rng.getCell(i,1).setBackground('#fcaba4');
          rng.getCell(i,2).setValue(errors.join(';'));
          valid = valid && false;
        }
        else
        {
          valid = valid && true;
        }
        i++;
      });
      
      return valid;
    }
    else
    {
      console.log('no required validations...');
      return true;
    }

    
  }
  catch(ex)
  {
    console.log(ex);
  }
  
  return valid;
}

function setAllPending_(rng,startRow,endRow)
{
  try
  {
    console.log('setAllPending');
    
    for(i = startRow; i <= endRow; i++)
    {
      rng.getCell(i,1).setValue('pending');
      rng.getCell(i,1).setBackground('#fcfcfc');
      rng.getCell(i,2).setValue('');
    }
  }
  catch(ex)
  {
    console.log(ex);
  }
}

function checkForJobStatus_(currentSheetId){
  
  try{
    console.log('checkForJobStatus');
    var options = getParseOptions_();
      //var result = UrlFetchApp.fetch(apiUrl + '/audit/status/' +   , options);
    var job = getAuditJobForCurrentSheet_(currentSheetId);
    if(job != null){
      var result = UrlFetchApp.fetch(getApiUrl_() + '/audit/status/' + job.audit.id , options);
      var statusCode = result.getResponseCode();
      console.log('Response Status Code = ' + statusCode);
      if(!handleResponse_(statusCode))
      {
         return;
      }
      else
      {
        if(statusCode === 200){
          var o  = Utilities.jsonParse(result.getContentText());
          if(o.success) {
            console.log('got audit...');
            if(o.audit.status == 'Pending')
            {
              Utilities.sleep(15000);
              checkForJobStatus_(currentSheetId);
            }
            else
            {
              if(o.audit.status == 'Completed') {                
                if(o.audit.output && o.audit.output.length > 0)
                {
                  var i = job.startRow;
                  var nRange = getNamedRange_ByName_(job.rangeName);
                  var rng = nRange.getRange();
                  o.audit.output.forEach(function(r){            
                    rng.getCell(i,1).setValue(r.status);
                    if(r.status == 'error' || r.status == 'invalid')
                    {
                      rng.getCell(i,2).setValue(r.message);
                      rng.getCell(i,1).setBackground('#fcaba4');
                    }
                    else
                    {
                      rng.getCell(i,1).setBackground('#c3fcdb');
                      rng.getCell(i,2).setValue('');
                    }            
                    i++;
                  });
                } 
                console.log('deleting audit job for current sheet...');
                deleteAuditJobForSheet_(currentSheetId);
              }
              else
              {
                SpreadsheetApp.getUi().alert('An error occured while uploading data. No rows were uploaded.');
              }
            }
          }
          else
          {
            console.log('is not success...');
            return;
          }
        }
      }
    }  
    else
    {
      console.log('job record is null...');
      return;
    }
  }
  catch(ex){
    console.log(ex);
  }
}

function uploadData(mode)
{
  try
  {
    console.log("uploadData...");    
    if(checkForExpiration_()){
      showUpgradePage_();
      return;
    }
    
    var user= currentUser_();
    if(user && user.planName === 'Starter')
    {
      SpreadsheetApp.getUi().alert('Templates are available only in Pro, Premium and Enterprise plans. Please consider upgrading your plan.');
      return;
    }    
    
    var startRow = -1;
    var endRow = -1;
    
    var namedRange = getNamedRangeForCurrentSheet_();
    if(namedRange == null)
    {
      console.log("cannot find namedRange for current cell...");
      SpreadsheetApp.getUi().alert("This sheet does not contain any template. Please select a sheet containing the template to upload.");
      return;
    }
    var tableName = namedRange.getName();
    
    var rng = namedRange.getRange();
   
    
    var template = getTmplForTable_(tableName);
   
    
    if(template == null)
    {
      console.log("cannot find template for current cell...");
      SpreadsheetApp.getUi().alert("This sheet does not contain any template. Please select a sheet containing the template to upload.");
      return;
    }
    
    if(!!!template.partialUpload && mode == 'Selection')
    {
      SpreadsheetApp.getUi().alert("Selected data upload is not enabled for this template. You can only upload all data.");
      return;
    }
    
    var data = getTemplateData_(tableName,mode == 'Selection');
    console.log('after getting template data..');
    //console.log(data);
    
    if(data == null)
      return;
    
    if(data != null && data.rows.length == 0)
    {
      console.log('data is null..returning..');
      SpreadsheetApp.getUi().alert("There is no data to upload!.");
      return;
    }
    
    if(mode == 'Selection')
    {
      startRow = SpreadsheetApp.getActiveRange().getRow();
      endRow = startRow + SpreadsheetApp.getActiveRange().getNumRows() - 1;
    }
    else
    {
      startRow = rng.getRow() + 1;
      endRow = getFirstEmptyRowWholeRow_() - 1;
    }
    
    //var batchSize = template.batchSize;
    
    
    //if(batchSize == undefined || batchSize == null || batchSize == 0 || batchSize == '')
    //  batchSize = data.rows.length;
    
    //console.log('batch Size = ' + batchSize);
    
    //var chunks = data.rows.chunk(batchSize);
    
    //var batch;
    //for(var batch= 1; batch <= chunks.length;batch++)
    //{
       console.log('startRow = ' + startRow);
      endRow = startRow + data.rows.length - 1;
      console.log('endRow = ' + endRow);
    
      setAllPending_(rng,startRow,endRow);
      SpreadsheetApp.flush();
    
      var options = getParseOptions_();
      options.payload = JSON.stringify({template: template.key, data: {rows: data.rows, cols: data.cols}});
      //console.log(options.payload);
      options["method"] = "post";
      //var result = UrlFetchApp.fetch(apiUrl + '/templates/upload?src=gsheets' , options);
      var result = UrlFetchApp.fetch(getApiUrl_() + '/templates/upload?src=gsheets' , options);
      var statusCode = result.getResponseCode();
      console.log('Response Status Code = ' + statusCode);
      if(!handleResponse_(statusCode))
      {        
        return;
      }
      else
      {
        if(statusCode == 200) {
          var o  = Utilities.jsonParse(result.getContentText());
         // console.log(o);
          if(o.success)
          {
            if(!!!o.background) {
              console.log('successful upload...');
              console.log(o.results);
              if(o.results && o.results.length > 0)
              {
                var i = startRow;
                o.results.forEach(function(r){            
                  rng.getCell(i,1).setValue(r.status);
                  if(r.status == 'error' || r.status == 'invalid')
                  {
                    rng.getCell(i,2).setValue(r.message);
                    rng.getCell(i,1).setBackground('#fcaba4');
                  }
                  else
                  {
                    rng.getCell(i,1).setBackground('#c3fcdb');
                    rng.getCell(i,2).setValue('');
                  }            
                  i++;
                });
              }
            }
            else
            {
              console.log('background upload started...');
              saveAuditJobForCurrentSheet_({rangeName: namedRange.getName(), startRow: startRow, endRow: endRow, audit:o.auditTrail});
              var currentSheetId = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId();
              checkForJobStatus_(currentSheetId);
            }
          }
          else
          {
            console.log('upload failure...');
            console.log(o.message);
            SpreadsheetApp.getUi().alert(o.message);
            SpreadsheetApp.getUi().alert('An error occurred while uploading data. Please try again later.');
            //break;
          }
          
        }
        else
        {
          console.log('Response Status Code = ' + statusCode);
          SpreadsheetApp.getUi().alert('Server Response: ' + statusCode + '. There is an error processing the data in the server. Please try again later!');
          //break;
        }
      }
    //}
  }
  catch(ex)
  {
    console.log(ex.message);
    SpreadsheetApp.getUi().alert(ex.message);
    SpreadsheetApp.getUi().alert("An error occurred while upload. Please try again later.");
  }
  
  return null;
}

function uploadSelectedData()
{
  try
  {
    console.log("uploadData...");
    uploadData('Selection');
  }
  catch(ex)
  {
    console.log(ex.message);
    SpreadsheetApp.getUi().alert("An error occurred while upload. Please try again later.");
  }
  
  return null;
}


function insertTemplate(template)
{
  try
  {
    console.log('insertTemplate...');
    var firstEmptyRow = SpreadsheetApp.getActiveSpreadsheet().getLastRow() + 1;
    if(firstEmptyRow > 1)
    {
      console.log("current sheet is not empty...");
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(generateSheetNameForQuery_(template.templateName));
    }
    
    var visibleCols = template.selectedCols.filter(function(item){return item.visible;});
    
    numCols= visibleCols.length + 2;
    numRows = 5000;
    
    var ssActive = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var myRange  = ssActive.getRange(1,1, numRows, numCols);
    console.log("after getting range"); 
    myRange.setBackground('#fcfcfc');
    var tblName = generateTblNameForQuery_(template.templateName);
    console.log("Tablename = " + tblName);
    createNamedRange_(tblName,myRange);
    console.log('after creating nm...');
    saveTableInfoForTmpl_(tblName,template);
    console.log("After inserting data..."); 
    var hdrRange = ssActive.getRange(1,1,1,numCols);
    hdrRange.setNumberFormat('@STRING@');
    var colHdrs = [];
    colHdrs.push('Status');
    colHdrs.push('Message');
    visibleCols.forEach(function(col) {
      colHdrs.push(col.caption);
    });
    
    console.log(colHdrs);
    hdrRange.setValues([colHdrs]);
    console.log('after setting headers...');
    
  }
  catch(ex)
  {
    if(ex.message == 'This action would increase the number of cells in the workbook above the limit of 2000000 cells.')
      SpreadsheetApp.getUi().alert('Google Sheets allows only 2000000 cells in a workbook. Please delete some unwanted sheets and try again.');
    else{
      SpreadsheetApp.getUi().alert('An error occured while rendering the template. Please try again later.');
      console.log(ex.message);
    }
    
  }
}

/*====================================================================================================================================*
  ImportJSON by Trevor Lohrbeer (@FastFedora)
  ====================================================================================================================================
  Version:      1.1
  Project Page: http://blog.fastfedora.com/projects/import-json
  Copyright:    (c) 2012 by Trevor Lohrbeer
  License:      GNU General Public License, version 3 (GPL-3.0) 
                http://www.opensource.org/licenses/gpl-3.0.html
  ------------------------------------------------------------------------------------------------------------------------------------
  A library for importing JSON feeds into Google spreadsheets. Functions include:
     ImportJSON            For use by end users to import a JSON feed from a URL 
     ImportJSONAdvanced    For use by script developers to easily extend the functionality of this library
  Future enhancements may include:
   - Support for a real XPath like syntax similar to ImportXML for the query parameter
   - Support for OAuth authenticated APIs
  Or feel free to write these and add on to the library yourself!
  ------------------------------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.1    Added support for the noHeaders option
  1.0    Initial release
 *====================================================================================================================================*/
/**
 * Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
 * a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
 * the JSON feed. The remaining rows contain the data. 
 * 
 * By default, data gets transformed so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *      of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    noHeaders:     Don't include headers, only the data
 *    debugLocation: Prepend each value with the row & column it belongs in
 *
 * For example:
 *
 *   =ImportJSON_("http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", "/feed/entry/title,/feed/entry/content",
 *               "noInherit,noTruncate,rawHeaders")
 * 
 * @param {url} the URL to a public JSON feed
 * @param {query} a comma-separated lists of paths to import. Any path starting with one of these paths gets imported.
 * @param {options} a comma-separated list of options that alter processing of the data
 *
 * @return a two-dimensional array containing the data, with the first row containing headers
 **/
function ImportJSON_(url, urlHeaders,noHeaders,query, options) {
  console.log('ImportJSON: ' + url);
  var result = UrlFetchApp.fetch(url,urlHeaders);
  var respCode = result.getResponseCode();
  console.log(result);
  if (!handleResponse_(respCode)) {   
    console.log('not a good response...');
    return null;
  }
  else
  {
    console.log('got response from server...');
  var object = JSON.parse(result.getContentText());
  console.log('object is...');
  console.log(object);
    
    return object;
  }
}

/**
 * An advanced version of ImportJSON designed to be easily extended by a script. This version cannot be called from within a 
 * spreadsheet.
 *
 * Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
 * a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
 * the JSON feed. The remaining rows contain the data. 
 *
 * Use the include and transformation functions to determine what to include in the import and how to transform the data after it is
 * imported. 
 *
 * For example:
 *
 *   =ImportJSON_("http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", 
 *               "/feed/entry",
 *                function (query, path) { return path.indexOf(query) == 0; },
 *                function (data, row, column) { data[row][column] = data[row][column].toString().substr(0, 100); } )
 *
 * In this example, the import function checks to see if the path to the data being imported starts with the query. The transform 
 * function takes the data and truncates it. For more robust versions of these functions, see the internal code of this library.
 *
 * @param {url}           the URL to a public JSON feed
 * @param {query}         the query passed to the include function
 * @param {options}       a comma-separated list of options that may alter processing of the data
 * @param {includeFunc}   a function with the signature func(query, path, options) that returns true if the data element at the given path
 *                        should be included or false otherwise. 
 * @param {transformFunc} a function with the signature func(data, row, column, options) where data is a 2-dimensional array of the data 
 *                        and row & column are the current row and column being processed. Any return value is ignored. Note that row 0 
 *                        contains the headers for the data, so test for row==0 to process headers only.
 *
 * @return a two-dimensional array containing the data, with the first row containing headers
 **/
function ImportJSONAdvanced_(url, headers,query, options, includeFunc, transformFunc) {
  console.log('ImportJSONAdvanced');
  var result = UrlFetchApp.fetch(url,headers);
  var respCode = result.getResponseCode();
  
  if (!handleResponse_(respCode)) {
    return null;
  }
  else
  {
  var object = JSON.parse(result.getContentText());
  //console.log('object is...');
  //console.log(object);
  
  return parseJSONObject_(object.data.rows, query, options, includeFunc, transformFunc);
  }
}

/** 
 * Encodes the given value to use within a URL.
 *
 * @param {value} the value to be encoded
 * 
 * @return the value encoded using URL percent-encoding
 */
function URLEncode_(value) {
  console.log('URLEncode');
  return encodeURIComponent(value.toString());  
}

/** 
 * Parses a JSON object and returns a two-dimensional array containing the data of that object.
 */
function parseJSONObject_(object, query, options, includeFunc, transformFunc) {
   console.log('parseJSONObject_');
  var headers = new Array();
  var data    = new Array();
  
  if (query && !Array.isArray(query) && query.toString().indexOf(",") != -1) {
    query = query.toString().split(",");
  }
  
  if (options) {
    options = options.toString().split(",");
  }
    
  parseData_(headers, data, "", 1, object, query, options, includeFunc);
  parseHeaders_(headers, data);
  transformData_(data, options, transformFunc);
  
  return hasOption_(options, "noHeaders") ? (data.length > 1 ? data.slice(1) : new Array()) : data;
}

/** 
 * Parses the data contained within the given value and inserts it into the data two-dimensional array starting at the rowIndex. 
 * If the data is to be inserted into a new column, a new header is added to the headers array. The value can be an object, 
 * array or scalar value.
 *
 * If the value is an object, it's properties are iterated through and passed back into this function with the name of each 
 * property extending the path. For instance, if the object contains the property "entry" and the path passed in was "/feed",
 * this function is called with the value of the entry property and the path "/feed/entry".
 *
 * If the value is an array containing other arrays or objects, each element in the array is passed into this function with 
 * the rowIndex incremeneted for each element.
 *
 * If the value is an array containing only scalar values, those values are joined together and inserted into the data array as 
 * a single value.
 *
 * If the value is a scalar, the value is inserted directly into the data array.
 */
function parseData_(headers, data, path, rowIndex, value, query, options, includeFunc) {
  
  //console.log('parseData_ for path:' + path + ' and value = ' + value);
  var dataInserted = false;
  
  if (value != null && isObject_(value)) {
    //console.log('value is object..');
    for (key in value) {
      if (parseData_(headers, data, path + "/" + key, rowIndex, value[key], query, options, includeFunc)) {
        dataInserted = true; 
      }
    }
  } else if (Array.isArray(value) && isObjectArray_(value)) {
    console.log('value is object array..');
    for (var i = 0; i < value.length; i++) {
      if (parseData_(headers, data, path, rowIndex, value[i], query, options, includeFunc)) {
        dataInserted = true;
        rowIndex++;
      }
    }
  } else if (!includeFunc || includeFunc(query, path, options)) {
    // Handle arrays containing only scalar values
   // console.log('value is scalar..');
    if (Array.isArray(value)) {
      value = value.join(); 
    }
    
    // Insert new row if one doesn't already exist
    if (!data[rowIndex]) {
      //console.log("adding new Row " + rowIndex);
      data[rowIndex] = new Array();
    }
    
    // Add a new header if one doesn't exist
    if (!headers[path] && headers[path] != 0) {
      headers[path] = Object.keys(headers).length;
      //console.log("Adding header: " + headers[path]);
    }
    
    // Insert the data
    data[rowIndex][headers[path]] = value;
    dataInserted = true;
  }
  
  return dataInserted;
}

/** 
 * Parses the headers array and inserts it into the first row of the data array.
 */
function parseHeaders_(headers, data) {
  
  //console.log('parseHeaders_');
  data[0] = new Array();

  for (key in headers) {
    data[0][headers[key]] = key;
  }
}

/** 
 * Applies the transform function for each element in the data array, going through each column of each row.
 */
function transformData_(data, options, transformFunc) {
   //console.log('transformData_');
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      transformFunc(data, i, j, options);
    }
  }
}

/** 
 * Returns true if the given test value is an object; false otherwise.
 */
function isObject_(test) {
  //console.log('isObject_');
  return  Object.prototype.toString.call(test) === '[object Object]';
}

/** 
 * Returns true if the given test value is an array containing at least one object; false otherwise.
 */
function isObjectArray_(test) {
  //console.log('isObjectArray_');
  for (var i = 0; i < test.length; i++) {
    if (isObject_(test[i])) {
      return true; 
    }
  }  

  return false;
}

/** 
 * Returns true if the given query applies to the given path. 
 */
function includeXPath_(query, path, options) {
   //console.log('includeXPath_');
  if (!query) {
    return true; 
  } else if (Array.isArray(query)) {
    for (var i = 0; i < query.length; i++) {
      if (applyXPathRule_(query[i], path, options)) {
        return true; 
      }
    }  
  } else {
    return applyXPathRule_(query, path, options);
  }
  
  return false; 
};

/** 
 * Returns true if the rule applies to the given path. 
 */
function applyXPathRule_(rule, path, options) {
   //console.log('applyXPathRule_');
  return path.indexOf(rule) == 0; 
}

/** 
 * By default, this function transforms the value at the given row & column so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *     of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Values in row 0 (headers) have slashes converted to spaces, common prefixes removed and the resulting text converted to title 
*      case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    debugLocation: Prepend each value with the row & column it belongs in
 */
function defaultTransform_(data, row, column, options) {
   //console.log('defaultTransform_');
  if (!data[row][column]) {
    if (row < 2 || hasOption_(options, "noInherit")) {
      data[row][column] = "";
    } else {
      data[row][column] = data[row-1][column];
    }
  } 

  if (!hasOption_(options, "rawHeaders") && row == 0) {
    if (column == 0 && data[row].length > 1) {
      removeCommonPrefixes_(data, row);  
    }
    
    data[row][column] = toTitleCase_(data[row][column].toString().replace(/[\/\_]/g, " "));
  }
  else
  {
    data[row][column] = data[row][column].toString().replace(/[\/]/g, "");
    //console.log("title = " + data[row][column]);
  }
  
  if (!hasOption_(options, "noTruncate") && data[row][column]) {
    data[row][column] = data[row][column].toString().substr(0, 256);
  }

  if (hasOption_(options, "debugLocation")) {
    data[row][column] = "[" + row + "," + column + "]" + data[row][column];
  }
}

/** 
 * If all the values in the given row share the same prefix, remove that prefix.
 */
function removeCommonPrefixes_(data, row) {
  var matchIndex = data[row][0].length;

  for (var i = 1; i < data[row].length; i++) {
    matchIndex = findEqualityEndpoint_(data[row][i-1], data[row][i], matchIndex);

    if (matchIndex == 0) {
      return;
    }
  }
  
  for (var i = 0; i < data[row].length; i++) {
    data[row][i] = data[row][i].substring(matchIndex, data[row][i].length);
  }
}

/** 
 * Locates the index where the two strings values stop being equal, stopping automatically at the stopAt index.
 */
function findEqualityEndpoint_(string1, string2, stopAt) {
   //console.log('findEqualityEndpoint_');
  if (!string1 || !string2) {
    return -1; 
  }
  
  var maxEndpoint = Math.min(stopAt, string1.length, string2.length);
  
  for (var i = 0; i < maxEndpoint; i++) {
    if (string1.charAt(i) != string2.charAt(i)) {
      return i;
    }
  }
  
  return maxEndpoint;
}
  

/** 
 * Converts the text to title case.
 */
function toTitleCase_(text) {
   //console.log('toTitleCase_');
  if (text == null) {
    return null;
  }
  
  return text.replace(/\w\S*/g, function(word) { return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); });
}

/** 
 * Returns true if the given set of options contains the given option.
 */
function hasOption_(options, option) {
  // console.log('hasOption_');
  return options && options.indexOf(option) >= 0;
}


function getTimeZoneList_(){
  
  return [
    {name:'Etc/GMT+12',desc:'Etc/GMT+12 (GMT-12:00)'},
{name:'Pacific/Midway',desc:'Pacific/Midway SST(GMT-11:00)'},
{name:'Pacific/Honolulu',desc:'Pacific/Honolulu HAST(GMT-10:00)'},
{name:'America/Anchorage',desc:'America/Anchorage AKST(GMT-09:00)'},
{name:'America/Los_Angeles',desc:'America/Los_Angeles PST(GMT-08:00)'},
{name:'America/Denver',desc:'America/Denver MST(GMT-07:00)'},
{name:'America/Chicago',desc:'America/Chicago CST(GMT-06:00)'},
{name:'America/New_York',desc:'America/New_York EST(GMT-05:00)'},
{name:'America/Santiago',desc:'America/Santiago CLT(GMT-04:00)'},
{name:'America/Sao_Paulo',desc:'America/Sao_Paulo BRT(GMT-03:00)'},
{name:'America/Noronha',desc:'America/Noronha FNT(GMT-02:00)'},
{name:'Atlantic/Cape_Verde',desc:'Atlantic/Cape_Verde CVT(GMT-01:00)'},
{name:'Europe/London',desc:'Europe/London GMT(GMT+00:00)'},
{name:'Europe/Amsterdam',desc:'Europe/Amsterdam CET(GMT+01:00)'},
{name:'Europe/Istanbul',desc:'Europe/Istanbul TRT(GMT+02:00)'},
{name: 'Europe/Moscow', desc: 'Europe/Moscow MSK(GMT+03:00)'},    
{name:'Asia/Kuwait',desc:'Asia/Kuwait AST(GMT+03:00)'},
{name:'Asia/Tehran',desc:'Asia/Tehran IRST(GMT+03:30)'},
{name:'Asia/Dubai',desc:'Asia/Dubai GST(GMT+04:00)'},
{name:'Asia/Kabul',desc:'Asia/Kabul AFT(GMT+04:30)'},
{name:'Asia/Karachi',desc:'Asia/Karachi PKT(GMT+05:00)'},
{name:'Asia/Kolkata',desc:'Asia/Kolkata IST(GMT+05:30)'},
{name:'Asia/Kathmandu',desc:'Asia/Kathmandu NPT(GMT+05:45)'},
{name:'Asia/Dhaka',desc:'Asia/Dhaka BKT(GMT+06:00)'},
{name:'Asia/Rangoon',desc:'Asia/Rangoon MMT(GMT+06:30)'},
{name:'Asia/Bangkok',desc:'Asia/Bangkok ICT(GMT+07:00)'},
{name:'Asia/Hong_Kong',desc:'Asia/Hong_Kong HKT(GMT+08:00)'},
{name:'Australia/Eucla',desc:'Australia/Eucla ACWST(GMT+08:45)'},
{name:'Asia/Tokyo',desc:'Asia/Tokyo JST(GMT+09:00)'},
{name:'Australia/Adelaide',desc:'Australia/Adelaide ACST(GMT+09:30)'},
{name:'Australia/Melbourne',desc:'Australia/Melbourne AEST(GMT+10:00)'},
{name:'Australia/Lord_Howe',desc:'Australia/Lord_Howe ACDT(GMT+10:30)'},
{name:'Pacific/Pohnpei',desc:'Pacific/Pohnpei AEDT(GMT+11:00)'},
{name:'Pacific/Norfolk',desc:'Pacific/Norfolk BST(GMT+11:30)'},
{name:'Pacific/Auckland',desc:'Pacific/Auckland NZST(GMT+12:00)'},
{name:'Pacific/Chatham',desc:'Pacific/Chatham CHAST(GMT+12:45)'},
{name:'Pacific/Enderbury',desc:'Pacific/Enderbury PHOT(GMT+13:00)'},
{name:'Pacific/Kiritimati',desc:'Pacific/Kiritimati WST(GMT+14:00)'}
    ];
  
}
