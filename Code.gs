/**
* @OnlyCurrentDoc
*/

/* Principles
 * Keep required permissions list as small as possible. In many cases it is possible to have all needed functionality with only one required permission: "View and manage spreadsheets that this application has been installed in"
 ** That is to avoid "You should avoid this app. Google is concerned this app may try to expose and exploit your private information. BACK TO SAFETY. Google hasn't reviewed this app yet and can't confirm it's authentic. Unverified apps may pose a threat to your personal data. Go to Aspire Budget Tools (unsafe)
 * No calls to any external services, as budget spreadsheets usually contain sensitive information, and calls to any external services may compromise the privacy.
 * Clear code, understandable by a novice, for anyone to be clear on the purpose of the code.
 * Clear code of bank message parsers, for the ability to debug and adapt instantly even for a novice when bank message format changes (and this happens from time to time).
 * For the sames reason code contains Quick Debug lines, ready to be used for a quick fixing when bank message format changes.
 * Clear code is more important than fast-running code, though it must not be too slow, as there are some limits: https://developers.google.com/apps-script/guides/services/quotas
 */
 
/* Quick Debug: basic overview of string parsing may be found here - https://www.w3schools.com/js/js_string_methods.asp */

/* GOOGLEFINANCE() sometimes fails to work and results in #N/A ("Google Finance internal error"). In this case: open Transactions > turn off all filters > F5 */

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Global variables
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

var availableToBudget = "Available to budget";
var accountTransfer = "↕️ Account Transfer";
var notificationEmail = "my~email@gmail.com";

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Spreadsheet extra menu functionality [original Aspire Budget 2.8 code]
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Aspire Budget')
      .addItem('Localize Dates and Currency', 'localize')
      .addToUi();
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Formatting functionality [original Aspire Budget 2.8 code]
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function localize() {
  localizeDate();
  localizeCurrency();
}

function localizeDate() {
 // Localization Tools Tab
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Localization Tools');
 var cell = sheet.getRange("H9");
 var newDate = cell.getNumberFormat();
  
 // Net Worth Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Net Worth');
 cell = sheet.getRange("B21:B");
 cell.setNumberFormat(newDate); 
  
 // Transactions Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
 cell = sheet.getRange("B8:B");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("B5:C5");
 cell.setNumberFormat(newDate);
  
 // Balances Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Balances');
 cell = sheet.getRange("B8:B");
 cell.setNumberFormat(newDate);
  
 // Category Transfers Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Category Transfers');
 cell = sheet.getRange("B9:B");
 cell.setNumberFormat(newDate);
  
 // Category Reports Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Category Reports');
 cell = sheet.getRange("B40:B");
 cell.setNumberFormat(newDate);
 
 // Account Reports Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Account Reports');
 cell = sheet.getRange("B29:B");
 cell.setNumberFormat(newDate);
 
 // Trend Reports Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trend Reports');
 cell = sheet.getRange("B35:B");
 cell.setNumberFormat(newDate);
  
 // Calculations Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calculations');
 cell = sheet.getRange("A21:A50");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("K2");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("K4");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("P13:Q13");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("P17:Q17");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("T3:T14");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("AB3:AB26");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("AB38:AC38");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("AU3:AU15");
 cell.setNumberFormat(newDate);
  
 cell = sheet.getRange("AT17:AU17");
 cell.setNumberFormat(newDate);
}

function localizeCurrency(){

 // Localization Tools Tab
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Localization Tools');
 var cell = sheet.getRange("H5");
 var newCurr = cell.getNumberFormat();
  
 // Configuration Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configuration');
 cell = sheet.getRange("D9:E68");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("B5:F5");
 cell.setNumberFormat(newCurr);
  
 // Dashboard Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
 cell = sheet.getRange("I4:O63");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("C5:C6");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("C10:C39");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("S11:S14");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("S17:S46");
 cell.setNumberFormat(newCurr);
  
 // Net Worth Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Net Worth');
 cell = sheet.getRange("C21:C");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("J4");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("J7");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("J10");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("H17:J17");
 cell.setNumberFormat(newCurr);
  
 // Transactions Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
 cell = sheet.getRange("C8:C");
 cell.setNumberFormat(newCurr);

 cell = sheet.getRange("D8:D");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("G3:H3");
 cell.setNumberFormat(newCurr);
  
 // Balances Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Balances');
 cell = sheet.getRange("C8:C");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("D8:D");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("E3:G3");
 cell.setNumberFormat(newCurr);
  
 // Category Transfers Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Category Transfers');
 cell = sheet.getRange("C9:C");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("F3:F4");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("D6");
 cell.setNumberFormat(newCurr);
  
 // Category Reports Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Category Reports');
 cell = sheet.getRange("F32:G32");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("E36:G36");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("C40:C");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("D40:D");
 cell.setNumberFormat(newCurr);
  
 // Account Reports Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Account Reports');
 cell = sheet.getRange("F25:G25");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("C29:C");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("D29:D");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("B6");
 cell.setNumberFormat(newCurr);
  
 // Trend Reports Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trend Reports');
 cell = sheet.getRange("F31:G31");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("C35:C");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("D35:D");
 cell.setNumberFormat(newCurr);
  
 // Calculations Tab
 sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calculations');
 cell = sheet.getRange("I1:J61");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("U3:Y14");
 cell.setNumberFormat(newCurr);
 
 cell = sheet.getRange("AC3:AR26");
 cell.setNumberFormat(newCurr);
  
 cell = sheet.getRange("AV3:BA15");
 cell.setNumberFormat(newCurr);
}

function resetTransactionsFormatting() {
  // Localization Tools Tab
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Localization Tools');
  var cell = sheet.getRange("H5");
  var newCurr = cell.getNumberFormat();
  
  // Transactions Tab
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  sheet.getRange("C8:C").setNumberFormat(newCurr);
  sheet.getRange("H8:H").setHorizontalAlignment("center");
  sheet.getRange(8,1,sheet.getMaxRows(),sheet.getMaxColumns()).setFontFamily("Roboto").setFontSize("11");
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Quick debug functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function getFullErrorInfo(err) {
  var errInfo = "ERROR:\n"; 
  for (var prop in err)  {  
    errInfo += "  property: " + prop + "\n    value: ["+ err[prop] + "]\n"; 
  } 
  errInfo += "  toString(): " + " value: [" + err.toString() + "]"; 
  return errInfo;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Functionality of webhook-triggered processing of new messages with transaction info (when external application calls Google Apps Script "Deploy as web app" webhook)
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

// This is a function that runs when the web app receives a GET request
function doGet(e) {
  // Not implemented
  return HtmlService.createHtmlOutput("GET request received");
}

// This is a function that runs when the web app receives a POST request
function doPost(e) {
  try {
    // Logging of all incoming requests, for the purposes of quick debugging in case of bank message format changes
    var logSheet = SpreadsheetApp.getActive().getSheetByName("Log");
    var logSheetRow = Math.max(logSheet.getLastRow(), 1) + 1;
    logSheet.insertRowAfter(logSheetRow - 1);
    logSheet.getRange(logSheetRow, 1).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"));
    logSheet.getRange(logSheetRow, 2).setValue(e.queryString);
    logSheet.getRange(logSheetRow, 3).setValue(e.parameter);
    logSheet.getRange(logSheetRow, 4).setValue(e.parameters);
    logSheet.getRange(logSheetRow, 5).setValue(e.postData.contents);
    
    // Quick Debug: logging to Debug sheet
    /*var debugSheet = SpreadsheetApp.getActive().getSheetByName("Debug");
    var debugSheetRow = Math.max(debugSheet.getLastRow(), 1) + 1;
    var debugSheetColumn = 1;
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")); SpreadsheetApp.flush();
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(typeof e.parameters); SpreadsheetApp.flush();
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(typeof e.parameters.evtprm1); SpreadsheetApp.flush();
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(typeof e.parameters.evtprm2); SpreadsheetApp.flush();
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(typeof e.parameters.evtprm3); SpreadsheetApp.flush();
    
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(typeof e.parameter); SpreadsheetApp.flush();
    if ('evtprm1' in e.parameters) {
      debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(e.parameters['evtprm1'][0]); SpreadsheetApp.flush();
    }
    if ('evtprm2' in e.parameters) {
      debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(e.parameters['evtprm2'][0]); SpreadsheetApp.flush();
    }*/
    
    if ('ipn_track_id' in e.parameters) {
      var transaction = parsePayPalIPN(e.parameters);
    } else if ('evtprm1' in e.parameters) {
      if (e.parameters['evtprm1'][0] == 'ru.sberbankmobile') {
        var transaction = parseSberbankPush(e.parameters.evtprm2[0], e.parameters.evtprm3[0]);
      }
    } else if (e.postData.contents != null) {
      var data = JSON.parse(e.postData.contents);
      if ((typeof data !== 'undefined') && (typeof data.bank !== 'undefined')) {
        if (data.bank == "sberbank") {
          var transaction = parseSberbankSMS(data.message, Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"));
        } else if (data.bank == "citibank") {
          var transaction = parseCitibankEmail(data.subject, data.message, Utilities.formatDate(data.message.getDate(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"));
        } else if (data.bank == "paypal") {
          var transaction = parsePayPalEmail(data.subject, data.message, Utilities.formatDate(data.message.getDate(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"));
        }
      }
    }

    // Adding transaction
    if (transaction !== null) {
      if (transaction.hasOwnProperty("error")) {
        switch (transaction.error) {
          case "filtered":
            logSheet.getRange(logSheetRow, 6).setValue("No");
            logSheet.getRange(logSheetRow, 7).setValue("OK");
            break;
          case "unknown format":
            logSheet.getRange(logSheetRow, 6).setValue("No");
            logSheet.getRange(logSheetRow, 7).setValue("Not OK");
            MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared in POST", messageSubject + "\n" + messageBody);
            break;
          default:
            MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared in POST", messageSubject + "\n" + messageBody);
        }
      } else {
        if (addTransaction(transaction)) {
          logSheet.getRange(logSheetRow, 6).setValue("Yes");
        } else {
          logSheet.getRange(logSheetRow, 6).setValue("No");
        }
        logSheet.getRange(logSheetRow, 7).setValue("OK");
      }
    } else {
      MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared in POST", messageSubject + "\n" + messageBody);
    }

    SpreadsheetApp.flush();
  } catch (error) {
    var errorSheet = SpreadsheetApp.getActive().getSheetByName("Log");
    var errorSheetRow = Math.max(errorSheet.getLastRow(), 1) + 1;
    errorSheet.insertRowAfter(errorSheetRow - 1);
    errorSheet.getRange(errorSheetRow, 1).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"));
    var fullErrorInfo = getFullErrorInfo(error);
    errorSheet.getRange(errorSheetRow, 2).setValue(fullErrorInfo);
    SpreadsheetApp.flush();
    MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared in POST", fullErrorInfo);
  }
  
  return HtmlService.createHtmlOutput("POST request received");
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Functionality of active processing of new messages with transactions (by calling function process() by Google Apps Script time-based trigger)
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function process() {
  checkGmail();
  //checkStatementFiles(); // TODO: process monthly bank statements
}

// EXTRA PERMISSIONS ARE REQUESTED BY THIS FUNCTION. This function calls GmailApp methods, and therefore requests the following permission: "Read, compose, send, and permanently delete all your email from Gmail"
function checkGmail() {
  try {
    var logSheet = SpreadsheetApp.getActive().getSheetByName("Log");
    
    // Quick Debug: logging to Debug sheet
    /*var debugSheet = SpreadsheetApp.getActive().getSheetByName("Debug");
    var debugSheetRow = Math.max(debugSheet.getLastRow(), 1) + 1;
    var debugSheetColumn = 1;
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")); SpreadsheetApp.flush();
    debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(e.parameters.evtprm3); SpreadsheetApp.flush();*/
    
    // Note: all phrases must be in lower case, as comparison below goes with s.toLowerCase()
    var searchQueries = {
      'PayPal': {'emails': ['service@paypal.com'], 'phrases': [], 'subjectPhrases': ['receipt', 'payment', 'payout', '"got money"', '"transferring money"', '"request has been approved"']},
      'YandexMoney': {'emails': ['inform@money.yandex.ru'], 'phrases': [], 'subjectPhrases': []},
      'Citibank': {'emails': ['citialerts.russia@citi.com'], 'phrases': [], 'subjectPhrases': []},
      'Sberbank': {'emails': ['900@unknown.email'], 'phrases': ['зачисление', 'перевод', 'выдача', 'покупка', 'оплата', '"мобильный банк"'], 'subjectPhrases': []}
    };

    for (var key in searchQueries) {
      // Form search queries
      var searchQuery = '-is:starred '; // search only for not starred emails
      var emailCount = searchQueries[key].emails.length;
      if (emailCount > 0) {
        searchQuery += 'from:(' + searchQueries[key].emails[0];
        for (var i = 1; i < emailCount; i++) {
          searchQuery += ' OR ' + searchQueries[key].emails[i];
        }
        searchQuery += ') ';
      }
      var phraseCount = searchQueries[key].phrases.length;
      if (phraseCount > 0) {
        searchQuery += '(' + searchQueries[key].phrases[0];
        for (var i = 1; i < phraseCount; i++) {
          searchQuery += ' OR ' + searchQueries[key].phrases[i];
        }
        searchQuery += ') ';
      }
      var subjectPhraseCount = searchQueries[key].subjectPhrases.length;
      if (subjectPhraseCount > 0) {
        searchQuery += 'subject:(' + searchQueries[key].subjectPhrases[0];
        for (var i = 1; i < subjectPhraseCount; i++) {
          searchQuery += ' OR ' + searchQueries[key].subjectPhrases[i];
        }
        searchQuery += ') ';
      }

      var threads = GmailApp.search(searchQuery);
      var messages = GmailApp.getMessagesForThreads(threads);
      for (var i = 0; i < messages.length; i++) {
        for (var j = 0; j < messages[i].length; j++) {
          var message = messages[i][j];
          
          // Check if the message meet the criteria of search query
          // Though we have checked for this in search query for threads, a thread may have some other messages and we need to filter them out
          if (message.isStarred()) continue;
          var messageFrom = message.getFrom();
          if (messageFrom.indexOf("<") != -1) {
            messageFrom = messageFrom.slice(messageFrom.indexOf("<") + 1, messageFrom.indexOf(">"));
          }
          var filterOut = true;
          searchQueries[key].emails.forEach(function(email) {
            if (email == messageFrom) {
              filterOut = false;
              return;
            }
          });
          if (filterOut) continue;
          
          var messageSubject = message.getSubject();
          var messageBody = message.getBody();
          
          // Logging of all found emails, for the purposes of quick debugging in case of bank message format changes
          var logSheetRow = Math.max(logSheet.getLastRow(), 1) + 1;
          logSheet.insertRowAfter(logSheetRow - 1);
          logSheet.getRange(logSheetRow, 1).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"));
          logSheet.getRange(logSheetRow, 2).setValue(messageSubject);
          logSheet.getRange(logSheetRow, 3).setValue(messageBody.replace(/\n/g, "\\n"));
          
          if (key == 'PayPal') {
            var transaction = parsePayPalEmail(messageSubject, messageBody, Utilities.formatDate(message.getDate(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"));
          } else if (key == 'YandexMoney') {
            var transaction = parseYandexMoneyEmail(messageSubject, messageBody, Utilities.formatDate(message.getDate(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"));
          } else if (key == 'Citibank') {
            var transaction = parseCitibankEmail(messageSubject, messageBody, Utilities.formatDate(message.getDate(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"));
          } else if (key == 'Sberbank') {
            var transaction = parseSberbankSMS(messageBody, Utilities.formatDate(message.getDate(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"));
          }
          
          // Adding transaction
          if (transaction !== null) {
            if (transaction.hasOwnProperty("error")) {
              switch (transaction.error) {
                case "filtered":
                  logSheet.getRange(logSheetRow, 6).setValue("No");
                  logSheet.getRange(logSheetRow, 7).setValue("OK");
                  break;
                case "unknown format":
                  logSheet.getRange(logSheetRow, 6).setValue("No");
                  logSheet.getRange(logSheetRow, 7).setValue("Not OK");
                  MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared", messageSubject + "\n" + messageBody);
                  break;
                default:
                  MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared", messageSubject + "\n" + messageBody);
              }
            } else {
              if (addTransaction(transaction)) {
                logSheet.getRange(logSheetRow, 6).setValue("Yes");
              } else {
                logSheet.getRange(logSheetRow, 6).setValue("No");
              }
              logSheet.getRange(logSheetRow, 7).setValue("OK");
            }
          } else {
            MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared", messageSubject + "\n" + messageBody);
          }
          GmailApp.starMessage(message);
          message.star().refresh();
          
          logSheetRow = logSheetRow + 1;
        }
      }
    }

    SpreadsheetApp.flush();
  } catch (error) {
    var errorSheet = SpreadsheetApp.getActive().getSheetByName("Log");
    var errorSheetRow = Math.max(errorSheet.getLastRow(), 1) + 1;
    errorSheet.insertRowAfter(errorSheetRow - 1);
    errorSheet.getRange(errorSheetRow, 1).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"));
    var fullErrorInfo = getFullErrorInfo(error);
    errorSheet.getRange(errorSheetRow, 2).setValue(fullErrorInfo);
    SpreadsheetApp.flush();
    MailApp.sendEmail(notificationEmail, "Transaction message of unknown format has appeared", fullErrorInfo);
  }
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Generic transaction adding functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function addTransaction(transaction) {
  // Quick Debug: examples of input variables
  //var bank = "citibank";
  //var bank = "sberbank";
  //var message = "Dear Customer, The following transaction has been charged to credit card number **7001: Amount: 100.00 RUB Point of sale: CITY-LINK.INFO         My Date: 11/04/2019 Available limit: 92,606.18 RUB";
  //var message = "\r\nDear Customer,\r\n\r\nYou have received a payment to your credit card:\r\nCard number **7001\r\nAmount: 15,000.00 RUB \r\nAvailable limit: 98,635.94 RUB\r\n \r\nAvailable cash limit: 98,635.94 RUB  \r\n\r\nIf you didn\u0432\u0402\u2122t authorize this transaction, please call CitiPhone immediately on +7(495)775-75-75 in Moscow, +7(812)336-75-75 in St. Petersburg or 8(800)700-38-38 elsewhere in Russia.\r\n\r\nInformation on how to dispute a charge to your account is available by visiting the \u0432\u0402\u045aFAQ\u0432\u0402\u045c page under \u0432\u0402\u045aContact Us\u0432\u0402\u045c at www.citibank.ru, or by clicking https://www.citibank.ru/russia/pdf/dispute_leaflet_rus.pdf  (how to dispute a charge to your debit or credit card) and ?         https://www.citibank.ru/russia/info/rus/pdf/disput_form_01-2017.pdf (Transaction Dispute Form).\r\n\r\nLearn more about Citibank Alerting Service on our website at www.citibank.ru.\r\n\r\nSincerely,\r\nAO Citibank\r\n\r\nPLEASE DO NOT REPLY TO THIS MESSAGE.\r\n  \r\nPlease let us know of any changes in your contact details by signing on to Citibank\u0412\u00ae Online and choosing \u0432\u0402\u045aContact Information\u0432\u0402\u045c under \u0432\u0402\u045aMy Profile\u0432\u0402\u045c. \r\nYou can set up Citibank Alerting Service preferences or opt out of receiving alerts by choosing \u0432\u0402\u045aCitibank Alerting Service\u0432\u0402\u045c under \u0432\u0402\u045aProducts & Services\u0432\u0402\u045c.";
  //var message = 'Перевод 33.03р от АННА АЛЕКСАНДРОВНА М.\n\n\nБаланс MAES5567: 61847.88р\n\n\nСообщение: "Возврат обратно"';

  // Quick Debug: logging with Logger (View menu > Logs / Ctrl+Enter)
  //Logger.log(dataContents);
  
  // Quick Debug: message box
  //Browser.msgBox(dataContents);
  
  // Quick Debug: logging to Debug sheet
  /*var debugSheet = SpreadsheetApp.getActive().getSheetByName("Debug");
  var debugSheetRow = Math.max(debugSheet.getLastRow(), 1) + 1;
  var debugSheetColumn = 1;
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")); SpreadsheetApp.flush();
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(parameters); SpreadsheetApp.flush();
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(typeof dataContents); SpreadsheetApp.flush();*/

  var transactionsSheet = SpreadsheetApp.getActive().getSheetByName("Transactions");
  var transactionsSheetRow = Math.max(transactionsSheet.getLastRow(), 1) + 1;
  transactionsSheet.insertRowAfter(transactionsSheetRow - 1);
  transactionsSheet.getRange(transactionsSheetRow, 2).setValue(transaction.date);
  transactionsSheet.getRange(transactionsSheetRow, 3).setValue(transaction.outflow);
  transactionsSheet.getRange(transactionsSheetRow, 4).setValue(transaction.inflow);
  transactionsSheet.getRange(transactionsSheetRow, 5).setValue(transaction.category);
  transactionsSheet.getRange(transactionsSheetRow, 6).setValue(transaction.account);
  transactionsSheet.getRange(transactionsSheetRow, 7).setValue(transaction.memo);
  if (typeof transaction.comment !== 'undefined') {
    transactionsSheet.getRange(transactionsSheetRow, 9).setValue(transaction.comment);
  }
  if (typeof transaction.keyword !== 'undefined') {
    transactionsSheet.getRange(transactionsSheetRow, 10).setValue(transaction.keyword);
  }
  SpreadsheetApp.flush();
  
  return true;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Bank and payment system transaction message processing functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

// Parsing of PayPal email
// Note: PayPal sends HTML that cannot be processed  with XmlService.parse(), it throws exceptions like 'The element type "style" must be terminated by the matching end-tag "</style>"'
function parsePayPalEmail(subject, message, date) {
  // Remove line breaks
  message = message.split("\n").join(" ");
  message = message.split("\r").join(" ");
  // Replace non-breaking spaces with regular spaces
  message = message.split(String.fromCharCode(160)).join(" ");
  message = message.split("&#160;").join(" ");
  // Remove duplicate spaces
  message = message.replace(/\s\s+/g, ' ');
  // Remove <head> part
  var plainBody = message.slice(message.indexOf('<body'), message.length);
  // Remove HTML markup
  var markupRegExp = new RegExp('<[^>]+>', 'gi');
  var plainBody = plainBody.replace(markupRegExp, ' ');
  plainBody = plainBody.replace(/\s\s+/g, ' ');
  
  plainBody = plainBody.split("&euro;").join("€");
  
  // Filter our irrelevant messages
  var tokenFilters = [
    "canceled your automatic payments",
    "Automatic payment profile set up"
  ];
  var tokenFilterFound = false;
  tokenFilters.forEach(function(tokenFilter) {
    if (message.indexOf(tokenFilter) != -1) {
      tokenFilterFound = true;
      return; // this is return only from this small anonymous (unnamed) function
    }
  });
  if (tokenFilterFound) {
    return {error:"filtered"};
  }
  
  if ((subject.indexOf("You've got money") != -1) || (plainBody.indexOf('Amount received:') != -1)) {
    // zhao​bao sent you 1,216.44 RUB.
    // Tilia has sent you 3.94 USD.
    // Shutterstock Images C.V. has sent you 11.01 USD.
    var tokenAmount = "sent you ";
    var tokenExtra = "Transaction Details";
    
    // Find transaction info in the message and put it to "Memo"
    var splitMessage = message.slice(0, message.indexOf(tokenAmount));
    var memoStart = splitMessage.lastIndexOf(">") + 1;
    var memo = message.slice(memoStart, message.indexOf(tokenExtra, memoStart)).trim();
    // Remove HTML markup from memo
    memo = memo.replace(markupRegExp, ' ');
    // Remove duplicate spaces from memo
    memo = memo.replace(/\s\s+/g, ' ');
    
    // Find amount in the message
    var amountStart = message.indexOf(tokenAmount) + tokenAmount.length;
    var amountEnd = message.indexOf("<", amountStart);
    var amount = message.slice(amountStart, amountEnd).trim();
    if (amount.substr(-1) == ".") {
      amount = amount.slice(0, amount.length - 1);
    }
    
    var category = availableToBudget;
  } else if (subject.indexOf("payment return") != -1) {
    // Notification of payment return for Transaction ID:39E224L525ND0174G • On 7 Dec 2019, you sent a payment to Ben Joel for 40.00 USD.\n\nThe funds have been returned to your account.
    var tokenMemo = "Hello ";
    var tokenPayment = " you sent a payment ";
    var tokenAmount = " for ";
    var tokenExtra = "Please contact";
    
    // Find transaction info in the message and put it to "Memo"
    var firstTokenMemoIndex = plainBody.indexOf(tokenMemo); // find the first entry of tokenMemo ("Hello ")
    var secondTokenMemoIndex = plainBody.indexOf(tokenMemo, firstTokenMemoIndex + tokenMemo.length); // find the second entry of tokenMemo ("Hello ")
    var commaIndex = plainBody.indexOf(",", secondTokenMemoIndex + tokenMemo.length); // find the entry of comma after the second entry of tokenMemo ("Hello ")
    var memo = plainBody.slice(commaIndex + 1, plainBody.indexOf(tokenExtra) - 1).trim();
    
    // Find amount in the message
    var tokenPaymentIndex = message.indexOf(tokenPayment);
    var amountStart = message.indexOf(tokenAmount, tokenPaymentIndex) + tokenAmount.length;
    var amountEnd = message.indexOf("<", amountStart);
    var amount = message.slice(amountStart, amountEnd).trim();
    if (amount.substr(-1) == ".") {
      amount = amount.slice(0, amount.length - 1);
    }
    
    var category = availableToBudget;
  } else if (subject.indexOf("We're transferring money to your bank") != -1) {
    // Total amount transferred	6 075,98 RUB	Bank account	ОАО "СБЕРБАНК РОССИИ" x-5552
    var tokenAmount = "Total amount transferred";
    var tokenBank = "Bank account";
    var tokenExtra = "Transaction ID";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenAmount) + tokenAmount.length, plainBody.indexOf(tokenExtra) - 1).trim();
    
    // Find amount in the message
    var amountStart = plainBody.indexOf(tokenAmount) + tokenAmount.length;
    var amountEnd = plainBody.indexOf(tokenBank, amountStart);
    var amount = plainBody.slice(amountStart, amountEnd).trim();
    
    var category = accountTransfer;
  } else if (subject.indexOf('Receipt for Your Payment') != -1) {
    // You sent a payment of $10,00 USD to Portal LLC
    var tokenAmount = "sent a payment of ";
    var tokenTo = " to ";
    var tokenExtra = "It may take a few moments";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenAmount) + tokenAmount.length, plainBody.indexOf(tokenExtra)).trim();
    
    // Find amount in the message
    var amountStart = plainBody.indexOf(tokenAmount) + tokenAmount.length;
    var amountEnd = plainBody.indexOf(tokenTo, amountStart);
    var amount = plainBody.slice(amountStart, amountEnd).trim();
    
    var keyword = memo.slice(memo.indexOf(tokenTo) + tokenTo.length).trim();
    var category = categorize(keyword);
  } else if (subject.indexOf('You have authorized a payment') != -1) {
    // You authorized a payment of $1 563,88 USD to Microsoft Corporation
    var tokenAmount = "authorized a payment of ";
    var tokenTo = " to ";
    var tokenExtra = "Your funds will be transferred";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenAmount) + tokenAmount.length, plainBody.indexOf(tokenExtra)).trim();
    
    // Find amount in the message
    var amountStart = plainBody.indexOf(tokenAmount) + tokenAmount.length;
    var amountEnd = plainBody.indexOf(tokenTo, amountStart);
    var amount = plainBody.slice(amountStart, amountEnd).trim();
    
    var keyword = memo.slice(memo.indexOf(tokenTo) + tokenTo.length).trim();
    var category = categorize(keyword);
  } else if (subject.indexOf('You sent a payment') != -1) {
    // You sent 46.00 USD to zhao​bao
    var tokenAmount = "You sent ";
    var tokenTo = " to ";
    var tokenExtra = "Transaction Details";
    var tokenTotalAmount = "You paid";
    var tokenExtra2 = "Help & Contact";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenAmount) + tokenAmount.length, plainBody.indexOf(tokenExtra) - 1).trim();
    
    // Find amount in the message
    var amountStart = plainBody.indexOf(tokenTotalAmount) + tokenTotalAmount.length;
    var amountEnd = plainBody.indexOf(tokenExtra2, amountStart);
    var amount = plainBody.slice(amountStart, amountEnd).trim();
    
    var keyword = memo.slice(memo.indexOf(tokenTo) + tokenTo.length).trim();
    var category = categorize(keyword);
  } else if ((subject.indexOf('Your payment') != -1) && (subject.indexOf('has been processed') != -1)) {
    // This email confirms that you have paid €10,12 EUR from your PayPal balance to Uber BV using PayPal.
    var tokenAmount = "you have paid ";
    var tokenFrom = " from ";
    var tokenTo = " to ";
    var tokenExtra = "using PayPal";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenAmount) + tokenAmount.length, plainBody.indexOf(tokenExtra)).trim();
    
    // Find amount in the message
    var amountStart = plainBody.indexOf(tokenAmount) + tokenAmount.length;
    var amountEnd = plainBody.indexOf(tokenFrom, amountStart);
    var amount = plainBody.slice(amountStart, amountEnd).trim();
    
    var keyword = memo.slice(memo.indexOf(tokenTo) + tokenTo.length).trim();
    var category = categorize(keyword);
  } else {
    return {error:"unknown format"};
  }

  // Erase leading currency symbol if present, as it duplicates currency code (examples in PayPal e-mails: "$10,00 USD", "$25,00 CAD", "€13,98 EUR", "139,00 RUB") 
  if (isNaN(amount.slice(0, 1))) {
    amount = amount.slice(1, amount.length);
  }
  if (amount.indexOf(".") != -1) {
    // Convert from X,XXX,XXX.XX to XXXXXXX.XX
    amount = amount.split(",").join("");
    // Convert from XXXXXXX.XX to XXXXXXX,XX
    amount = amount.replace(".", ",");
    // Apply currency conversion if needed
  } else {
    var currency = amount.substr(-4);
    var amountWithoutCurrency = amount.slice(0, amount.length - 4);
    // Convert from X XXX XXX,XX to XXXXXXX,XX
    amount = amountWithoutCurrency.split(" ").join("") + currency;
  }
  // Apply currency conversion if needed
  amount = parseAmountWithCurrency(amount, date);

  // Define outflow / inflow
  if (category == availableToBudget) {
    var outflow = "";
    var inflow = amount;
  } else {
    var outflow = amount;
    var inflow = "";
  }

  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:"💰 Alexander’s PayPal", memo:memo, keyword:keyword};
  } else {
    return null;
  }
}

// Parsing of PayPal IPN request
// Though as PayPal IPN reports only inflow transactions, but does not report outflow transactions, parsePayPalEmail() is recommended to be used instead
function parsePayPalIPN(parameters) {
  // Not implemented
  return null;
}

// Parsing of Yandex.Money email
// Note: Yandex.Money sends HTML that cannot be processed with XmlService.parse(), it throws exceptions like 'The markup in the document following the root element must be well-formed'
function parseYandexMoneyEmail(subject, message, date) {
  // Remove line breaks
  message = message.split("\n").join(" ");
  message = message.split("\r").join(" ");
  // Replace non-breaking spaces with regular spaces
  message = message.split(String.fromCharCode(160)).join(" ");
  message = message.split("&#160;").join(" ");
  // Remove duplicate spaces
  message = message.replace(/\s\s+/g, ' ');
  
  // Unify currency text
  var tokenCurrency = "RUB";
  message = message.split("руб.").join(tokenCurrency);
  
  // Remove HTML markup
  var xmlRegExp = new RegExp('<[^>]+>', 'gi');
  var plainBody = message.replace(xmlRegExp, ' ');
  plainBody = plainBody.replace(/\s\s+/g, ' ');
  
  // Filter our irrelevant messages
  var tokenFilters = [
    "Информация о платеже", // Не обрабатываем, т. к. это не списание денег из Яндекс.Денег, это списание денег с других банковских карт и платёжных инструментов, проведенное через интерфейс Яндекс.Денег
    "Вы отправили перевод с карты на кошелек",
    "Вы заплатили с привязанной банковской карты",
    "Кэшбэк",
    "Доставка выписки"
  ];
  var tokenFilterFound = false;
  tokenFilters.forEach(function(tokenFilter) {
    if (message.indexOf(tokenFilter) != -1) {
      tokenFilterFound = true;
      return; // this is return only from this small anonymous (unnamed) function
    }
  });
  if (tokenFilterFound) {
    return {error:"filtered"};
  }
  
  if (((subject.indexOf('Ваш кошелек') != -1) || (subject.indexOf('Ваш счет') != -1)) && (subject.indexOf('пополнен') != -1)) {
    var tokenDate = "Дата и время";
    var tokenMemo = "Пополнение через";
    var tokenAmount = "Сумма";
    var tokenBalance = "Доступно";
    var tokenExtra = "Все детали";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenMemo) + tokenMemo.length, plainBody.indexOf(tokenAmount)).trim();
    
    // Find date in the message
    date = plainBody.slice(plainBody.indexOf(tokenDate) + tokenDate.length, plainBody.indexOf(tokenMemo)).trim();
    
    var category = availableToBudget;
  } else if (subject.indexOf("возврат по операции") != -1) {
    var tokenAmount = "Сумма возврата";
    var tokenMemo = "Где был платёж";
    var tokenBalance = "Доступно";
    var tokenExtra = "Запись об операции";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = subject.slice(subject.indexOf("возврат по операции"), subject.length) + " | " + plainBody.slice(plainBody.indexOf(tokenMemo) + tokenMemo.length, plainBody.indexOf(tokenBalance)).trim();
    
    var category = availableToBudget;
  } else if (subject.indexOf("Вернули на баланс") != -1) {
    var tokenAmount = "Зачислено";
    var tokenDate = "Дата и время";
    var tokenMemo = "Где был платёж";
    var tokenBalance = "Доступно";
    var tokenExtra = "Банковская карта";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = subject + " | " + plainBody.slice(plainBody.indexOf(tokenMemo), plainBody.indexOf(tokenBalance)).trim();
    
    // Find date in the message
    date = plainBody.slice(plainBody.indexOf(tokenDate) + tokenDate.length, plainBody.indexOf(tokenMemo)).trim();
    
    var category = availableToBudget;
  } else if (subject.indexOf('Вы заплатили из кошелька') != -1) {
    var tokenMemo = "Назначение платежа";
    var tokenDate = "Дата и время";
    var tokenAmount = "Списано";
    var tokenBalance = "Доступно";
    var tokenExtra = "Все детали";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenMemo) + tokenMemo.length, plainBody.indexOf(tokenDate)).trim();
    
    // Find date in the message
    date = plainBody.slice(plainBody.indexOf(tokenDate) + tokenDate.length, plainBody.indexOf(tokenAmount)).trim();
    
    var keyword = memo;
    var category = categorize(keyword);
  } else if (subject.indexOf('Вы заплатили с карты Яндекс.Денег') != -1) {
    var tokenMemo = "Назначение платежа";
    var tokenDate = "Дата и время";
    var tokenAmount = "Сколько списано";
    var tokenBalance = "Доступно";
    var tokenExtra = "Все детали";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = plainBody.slice(plainBody.indexOf(tokenMemo) + tokenMemo.length, plainBody.indexOf(tokenDate)).trim();
    
    // Find date in the message
    date = plainBody.slice(plainBody.indexOf(tokenDate) + tokenDate.length, plainBody.indexOf(tokenAmount)).trim();
    
    var keyword = memo;
    var category = categorize(keyword);
  } else if (subject.indexOf('Списали курсовую разницу') != -1) {
    var tokenAmount = "Списано";
    var tokenDate = "Дата и время";
    var tokenMemo = "Где был платёж";
    var tokenBalance = "Доступно";
    
    // Find transaction info in the message and put it to "Memo"
    var memo = subject + " | " + plainBody.slice(plainBody.indexOf(tokenMemo), plainBody.indexOf(tokenBalance)).trim();
    
    // Find date in the message
    date = plainBody.slice(plainBody.indexOf(tokenDate) + tokenDate.length, plainBody.indexOf(tokenMemo)).trim();
    
    var keyword = memo;
    var category = categorize(keyword);
  } else {
    return {error:"unknown format"};
  }

  // Find amount in the message
  var amountStart = plainBody.indexOf(tokenAmount) + tokenAmount.length;
  var amountEnd = plainBody.indexOf(tokenCurrency, amountStart);
  var amount = plainBody.slice(amountStart, amountEnd).trim();

  // Find balance in the message
  var balanceStart = plainBody.indexOf(tokenBalance) + tokenBalance.length;
  var balanceEnd = plainBody.indexOf(tokenCurrency, balanceStart);
  var balance = plainBody.slice(balanceStart, balanceEnd).trim();
    
  // Convert from X XXX XXX,XX to XXXXXXX,XX
  amount = amount.split(" ").join("");
  balance = balance.split(" ").join("");

  // Define outflow / inflow
  if (category == availableToBudget) {
    var outflow = "";
    var inflow = amount;
  } else {
    var outflow = amount;
    var inflow = "";
  }
  
  var account = "💰 Alexander’s Yandex.Money";
  var comment = isReconciled(account, Number(balance.split(",").join("."))) ? "Reconciled" : "Not reconciled";

  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:account, memo:memo, comment:comment, keyword:keyword};
  } else {
    return null;
  }
}

// EXTRA PERMISSIONS ARE REQUESTED BY THIS FUNCTION. This function calls UrlFetchApp.fetch('http://money.yandex.ru/oauth/authorize', options), and therefore requests the following permission: "Connect to an external service"
// First, personal client_id need to be generated: https://yandex.ru/dev/money/doc/dg/tasks/register-client-docpage/
function authorizeYandexMoney() {
  // Not implemented
  return null;
}

// Parsing of Sberbank SMS
function parseSberbankSMS(message, date) {
  // Quick Debug: examples of input variable
  //var message = 'Перевод 33.03р от АННА АЛЕКСАНДРОВНА М.\n\n\nБаланс MAES5567: 11847.88р\n\n\nСообщение: "Возврат обратно"';
  //var message = 'MAES5567 12:14 Зачисление зарплаты 50000р Баланс: 61847.50р';
  //var message = 'MAES5567 10:59 Зачисление аванса 50000р Баланс: 53747.50р';
  //var message = 'MAES5567 10:16 Зачисление пособия на детей 3277.45р Баланс: 11827.11р';
  //var message = 'MAES5567 10:06 зачисление 1999.85р Баланс: 16847.35р';
  //var message = 'MAES5567 02:50 зачисление под отчет 27408.93р';
  //var message = 'MAES5567 03:36 зачисление 1596.71р СО ВКЛАДА N*138019737008-1596.71RUR Баланс: 12775.49р';
  //var message = "MAES5567 12:05 Выдача 5000р ATM 95460011 Баланс: 12839.78р";
  //var message = "MAES5567 14:35 Покупка 298р STARBUCKS Баланс: 14848.78р";
  //var message = 'MAES5567 09:41 Оплата 3000р Баланс: 10080.29р';
  //var message = 'MAES5567 12:42 Оплата 500р NKO YANDEKS.DENGI Баланс: 93.35р';
  //var message = 'MAES5567 14:50 перевод 1400р Баланс: 13598.78р'; // Для перевода 1400р получателю АННА АЛЕКСАНДРОВНА М. на VISA4753 с карты MAES5567 отправьте код 62733 на 900.Комиссия не взимается
  //var message = 'MAES5567 02.12.19 мобильный банк за 02.12-01.01 30р Баланс: 16428.91р';
  // Сообщения по старому формату 2017:
  //var message = 'MAES4244 02.03.18 оплата Мобильного банка за 02/03/2018-01/04/2018 30р';
  //var message = 'MAES4244 05.03.18 11:21 списание 1640р';
  //var message = 'MAES4244 13.09.16 12:25 списание 30000р ATM 815226';
  //var message = 'MAES4244 13.09.16 12:25 отмена списания 30000р ATM 815226';
  //var message = 'MAES4244 26.09.17 14:12 оплата услуг 300р MTS OAO';
  //var message = 'MAES4244 22.12.17 11:51 выдача 2000р ATM 13021906';
  //var message = 'MAES4244 11.06.18 11:40 выдача 103USD с комиссией 100р 324 E 1ST ST';
  //var message = 'MAES4244: перевод 725р. на карту получателя АЛЕКСАНДР АЛЕКСАНДРОВИЧ Н. выполнен. Подробнее в выписке по карте http://sberbank.ru/sms/h2/';
  //var message = 'MAES4244 16.11.17 12:59 оплата 10500р';
  //var message = 'MAES4244 21.04.18 12:04 покупка 3118.08р OOO NOVYY IMPULS-50';
  //var message = 'Сбербанк Онлайн. АЛЕКСАНДР АЛЕКСАНДРОВИЧ Н. перевел(а) Вам 1140.00 RUB';
  // Сообщения по старому формату 2015:
  //var message = 'MAES1705: 02.03.15 оплата Мобильного банка за 02/03/2015-01/04/2015 на сумму 30.00р.';
  //var message = 'MAES1705: 16.02.15 16:46 операция зачисления на сумму 0.21р.';
  //var message = 'MAES1705: 18.11.14 12:42 операция перевода на сумму 2160.00р. OSB 5281 1522.';
  //var message = 'MAES1705: 27.02.15 12:23 операция списания на сумму 3500.00р. SBOL.';
  //var message = 'MAES1705: 26.12.14 11:37 выдача наличных на сумму 4000.00р. ATM 850726.';
  //var message = 'MAES1705: 31.03.15 08:24 оплата услуг на сумму 250.00р. MTS OAO (9119260960).';
  //var message = 'MAES1705: 29.03.14 13:32 оплата за выписку/запрос баланса на сумму 15.00 руб. ITT 852840 90400081 выполнена успешно. Доступно: 18570.93 руб.';
  //var message = '31.05.17 АННА АЛЕКСАНДРОВНА М. оплатил(а) Ваш телефон 9101221212 500р. Ваш Сбербанк!';
  //var message = 'MAES1705 22.06.15 14:43 взыскание/арест по требованию судебных органов 154.92р';
  // Сообщения по старому формату 2011:
  //var message = 'MAES2341; Popolnenie scheta; Uspeshno; Summa:5000.00RUR; ITT 852685 52812600; 04.10.11 12:41; Dostupno:16171.18RUR;';
  //var message = 'MAES2341; Perevod na karty ili vznos nalichnyh; Uspeshno; Summa:5000.00RUR; BANKOMAT 890442 7982 s karty 4252****3489; 24.11.11 14:25; Dostupno:14127.92RUR;';
  //var message = 'MAES2341; Vydacha nalichnyh; Uspeshno; Summa:6500.00RUR; BANKOMAT 814236 7970; 21.07.11 17:31; Dostupno:13681.18RUR;';
  //var message = 'MAES2341; Spisanie: perevod sredstv na karty; Uspeshno; Summa:3000.00RUR; SBOL; 08.08.11 12:08; Dostupno:12251.18RUR;';
  //var message = 'MAES2341; Pokupka; Uspeshno; Summa:2012.00RUR; ZOOMAGAZIN CHETYRE LAPY; 07.10.11 17:34; Dostupno:14059.18RUR;';
  //var message = 'MAES2341; Beznalichny perevod sredstv; Uspeshno; Summa:2023.65RUR; OSB 7970 0495; 14.11.11 19:04; Dostupno:14344.17RUR;';
  //var message = 'MAES2341; Oplata uslug; Uspeshno; Summa:50.00RUR; MTS OAO; 08.08.11 19:20; Dostupno:18631.18RUR;';
  //var message = 'MAES2341; Oplata uslug mobilnogo banka za period s 12/08/2011 po 11/09/2011; Uspeshno; Summa:30.00RUR; 12.08.11; Dostupno:18601.18RUR;';
  
  // Quick Debug: logging to Debug sheet
  /*var debugSheet = SpreadsheetApp.getActive().getSheetByName("Debug");
  var debugSheetRow = Math.max(debugSheet.getLastRow(), 1) + 1;
  var debugSheetColumn = 1;
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")); SpreadsheetApp.flush();
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(message); SpreadsheetApp.flush();*/

  var tokenTypesWithCategories = {
    " Зачисление": availableToBudget,
    " зачисление": availableToBudget,
    "Перевод": availableToBudget,
    " перевел(а) Вам": availableToBudget,
    " операция зачисления на сумму": availableToBudget,
    " отмена списания": availableToBudget,
    "Perevod na karty ili vznos nalichnyh; Uspeshno; Summa:": availableToBudget,
    "Popolnenie scheta; Uspeshno; Summa:": accountTransfer,
    " выдача наличных на сумму": accountTransfer,
    " выдача наличных": accountTransfer,
    " выдача": accountTransfer,
    " Выдача": accountTransfer,
    " Vydacha nalichnyh; Uspeshno; Summa:": accountTransfer,
    " Покупка": null,
    " покупка на сумму": "",
    " покупка": "",
    " мобильный банк за ": "Internet & Phone", // для унификации кода определения var amountStart оставляем в данной подстроке завершающий пробел, т. к. в сообщении от Сбербанка сумма транзакции идёт не сразу за данной подстрокой
    " оплата Мобильного банка за ": "Internet & Phone", // для унификации кода определения var amountStart оставляем в данной подстроке завершающий пробел, т. к. в сообщении от Сбербанка сумма транзакции идёт не сразу за данной подстрокой
    " оплата за выписку/запрос баланса на сумму": "Taxes & Fees",
    " взыскание/арест по требованию судебных органов": "Taxes & Fees",
    " оплата услуг на сумму": "",
    " оплата услуг": "",
    " Оплата": null,
    " оплата": "",
    " перевод": "",
    " списание": "",
    " операция списания на сумму": "",
    " операция перевода на сумму": "",
    " оплатил(а) Ваш телефон 9101221221": "",
    " Spisanie: perevod sredstv na karty; Uspeshno; Summa:": "",
    " Pokupka; Uspeshno; Summa:": "",
    " Beznalichny perevod sredstv; Uspeshno; Summa:": "",
    " Oplata uslug; Uspeshno; Summa:": "",
    " Oplata uslug mobilnogo banka za period s ": "Internet & Phone"
  };

  var tokenBalance = "Баланс";
  var tokenComment = "Сообщение: ";
  var tokenCommission = " с комиссией ";

  // Replace line breaks with spaces
  message = message.split("\n").join(" ");
  message = message.split("\r").join(" ");
  // Replace non-breaking spaces with regular spaces
  message = message.split(String.fromCharCode(160)).join(" ");
  //var messageLowerCase = message.toLowerCase();

  // Filter our irrelevant messages
  // Пароль для входа в Сбербанк Онлайн: 43436. НИКОМУ не сообщайте пароль.
  // Пароль для подтверждения платежа - 05218. Оплата 1500,00 RUB с карты **** 5567. Реквизиты: Регистратор R
  // СМС-код для входа в профиль: 0271
  // Для перевода 400р получателю АННА АЛЕКСАНДРОВНА М. на VISA2044 с карты MAES5567 отправьте код 71836 на 900.Комиссия не взимает
  // Операция не выполнена. Нет доступных карт для перевода или на картах недостаточно средств. Для проведения перевода пополните дебетовую карту и повторите попытку. Переводы с кредитных карт недоступн
  // Открытие брокерского  счёта. Пароль для подтверждения - 33564. Никому его не сообщайте.
  // Сбербанк Бизнес Онлайн. Код для входа 04327. Не сообщайте код никому.
  // Уважаемый клиент, карта VISA5572 активирована. Сбербанк
  // К карте VISA5572 подключен Мобильный банк без уведомлений об операциях. Если вы не подключали услугу, позвоните по номеру 900. Для подключения уведомлений об операциях отправьте СМС на номер 900 с текстом «Полный»
  // Семён Семёнович, мы подготовили для вас специальное предложение: 950 000р. в кредит по ставке 17.9% годовых на 60 мес., платёж 24 073 р./мес. Узнайте детали и оформите заявку у вашего персонального менеджера или самостоятельно в Сбербанк Онлайн: http://sberbank.ru/dl/cl/  ПАО Сбербанк
  // Семён Семёнович, теперь со своим персональным менеджером вы можете связаться по короткому номеру - 0440. Спасибо, что вы с нами! ПАО Сбербанк
  // Семён Семёнович, звоните своему персональному менеджеру по короткому номеру 0440: мы сразу будем знать, от кого звонок, и сможем помочь быстрее. Звонки доступны на территории России для абонентов Билайн, Мегафон, МТС и Теле2. Подробнее: www.sberbank.ru/t/0440  ПАО Сбербанк
  // СЕМЁН СЕМЁНОВИЧ, 15.07.2019 20:30:22 выполнена регистрация в приложении "Сбербанк Онлайн" для Android. Если вы не совершали операцию, позвоните по номеру 900
  // Семён Семёнович, оцените, пожалуйста, насколько Вы довольны обслуживанием в Сбербанк Премьер, пройдя короткий опрос по ссылке. Сбербанк https://opros.sberbank.ru/8z8brlc
  var tokenFilters = [
    "Совершен вход ",
    "СМС-код",
    "Пароль для",
    "пароль для",
    "Для перевода",
    "Не сообщайте",
    "не сообщайте",
    "не выполнена",
    "Семён Семёнович, ",
    "СЕМЁН СЕМЁНОВИЧ, ",
    "Семён Семёнович! ",
    "СЕМЁН СЕМЁНОВИЧ! ",
    "Сбербанк Бизнес Онлайн",
    "SBBOL"
  ];
  var tokenFilterFound = false;
  tokenFilters.forEach(function(tokenFilter) {
    if (message.indexOf(tokenFilter) != -1) {
      tokenFilterFound = true;
      return; // this is return only from this small anonymous (unnamed) function
    }
  });
  if (tokenFilterFound) {
    return {error:"filtered"};
  }
  
  var keyword = "";
  
  // Find transaction info in the message and put it to "Memo"
  var balanceStart = message.indexOf(tokenBalance);
  var commentStart = message.indexOf(tokenComment);
  if (balanceStart != -1) {
    var memo = message.slice(0, message.indexOf(tokenBalance) - 1).trim();
    if (commentStart != -1) {
      memo = memo + " " + message.slice(commentStart).trim();
    }
  } else {
    var memo = message;
  }
  
  // Find amount in the message, identify its type and determine category
  var category = "";
  for (var tokenType in tokenTypesWithCategories) {
    if (message.indexOf(tokenType) != -1) {
      var spaceIndex = message.indexOf(" ", message.indexOf(tokenType) + tokenType.length) + 1; // find the first whitespace after the token
      var messagePart = message.substr(spaceIndex);
      var digitIndex = messagePart.search(/\d/); // find the first digit in message part
      var amountStart = spaceIndex + digitIndex;
      var amountEnd = message.indexOf(" ", amountStart);
      // Legacy code: for parsing SMS of 2011
      //var amountStart = message.indexOf(tokenType) + tokenType.length;
      //var amountEnd = message.indexOf("RU", amountStart);

      category = tokenTypesWithCategories[tokenType];
      if (category == null) {
        keyword = message.slice(amountEnd, message.indexOf(tokenBalance)).trim();
        category = categorize(keyword);
      }
      break;
    }
  }

  if (date == null) {
    // Take current date
    date = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy");
  }

  var amount = message.slice(amountStart, amountEnd).trim();
  // Erase ending currency symbol
  if (amount.substr(-1) == "р") {
    amount = amount.slice(0, amount.length - 1).trim();
  } else if (amount.substr(-2) == "р.") {
    amount = amount.slice(0, amount.length - 2).trim();
  }
  // Convert from X,XXX,XXX.XX to XXXXXXX.XX
  amount = amount.split(",").join("");
  // Convert from XXXXXXX.XX to XXXXXXX,XX
  amount = amount.replace(".", ",");
  
  // Apply currency conversion if needed
  amount = parseAmountWithCurrency(amount, date);

  if (message.indexOf(tokenCommission) != -1) {
    var commissionAmountStart = message.indexOf(tokenCommission) + tokenCommission.length;
    var commissionAmountEnd = message.indexOf(" ", commissionAmountStart);
    var commissionAmount = message.slice(commissionAmountStart, commissionAmountEnd).trim();
    // Erase ending currency symbol
    if (commissionAmount.substr(-1) == "р") {
      commissionAmount = commissionAmount.slice(0, commissionAmount.length - 1).trim();
    } else if (commissionAmount.substr(-2) == "р.") {
      commissionAmount = commissionAmount.slice(0, commissionAmount.length - 2).trim();
    }
    // Convert from X,XXX,XXX.XX to XXXXXXX.XX
    commissionAmount = commissionAmount.split(",").join("");
    // Convert from XXXXXXX.XX to XXXXXXX,XX
    commissionAmount = commissionAmount.replace(".", ",");
    // Apply currency conversion if needed
    commissionAmount = parseAmountWithCurrency(commissionAmount, date);
    
    if (amount.slice(0, 1) != "=") {
      amount = "=" + amount;
    }
    amount = amount + "+" + commissionAmount;
  }

  // Define outflow / inflow
  if (category == availableToBudget) {
    var outflow = "";
    var inflow = amount;
  } else {
    var outflow = amount;
    var inflow = "";
  }
  
  var account = "";
  if (message.indexOf("MAES5567") != -1) {
    account = "💰 Alexander’s Sberbank Maestro";
  } else if (message.indexOf("ECMC1457") != -1) {
    account = "💰 Anna’s Sberbank MasterCard";
  }

  // TODO: implement parsing of balance amount in the message to perform check if the records are reconciled
  var balance = "0";
  var comment = isReconciled(account, Number(balance.split(",").join("."))) ? "Reconciled" : "Not reconciled";

  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:account, memo:memo, comment:comment, keyword:keyword};
  } else {
    return null;
  }
}

// Parsing of Sberbank push notification
// Though as Sberbank app does not do push notifications when device is offline, but sends SMS instead in this case, parseSberbankSMS() is recommended to be used instead
function parseSberbankPush(subject, message, date) {
  // Quick Debug: examples of input variable
  //var subject = 'Зачисление';
  //var message = '+1 999,85 ₽ - Баланс= 17 429,34 ₽ MAES •• 5567';
  //var subject = '💳 Зачисление зарплаты';
  //var message = '+50 000 ₽ - Баланс= 62 353,34 ₽ Maestro •• 5567';
  //var subject = '💳 Зачисление аванса';
  //var message = '+50 000 ₽ - Баланс= 55 421,34 ₽ Maestro •• 5567';
  //var subject = 'Зачисление Анна Александровна М.';
  //var message = '+250 ₽ - Баланс= 12 430,20 ₽ Maestro •• 5567 "Возврат обратно"';
  //var subject = 'Перевод Сбербанк Онлайн';
  //var message = '1,23 ₽ - Баланс= 12 103,87 ₽ Maestro •• 5567';
  //var subject = 'Выдача Сбербанк';
  //var message = '5 000 ₽ - Баланс= 7 438,87 ₽ Maestro •• 5567';
  //var subject = 'Покупка Ашан';
  //var message = '139,29 ₽ - Баланс= 9 714,87 ₽ Maestro •• 5567';
  //var subject = 'Покупка Нияма';
  //var message = '299 ₽ - Баланс= 8 571,87 ₽ Maestro •• 5567';
  //var subject = 'Оплата Сбербанк';
  //var message = '1 000 ₽ - Баланс= 12 742,87 ₽ Maestro •• 5567';
  //var subject = 'Оплата Мобильного банка';
  //var message = '60 ₽ - Баланс= 12 523,87 ₽ Maestro •• 5567';
  //var subject = 'Оплата Стрелка';
  //var message = '300 ₽ - Баланс= 11 842,87 ₽ Maestro •• 5567';
  //var subject = 'Оплата Автоплатёж МТС';
  //var message = '200 ₽ - Баланс= 14 924,87 ₽ Maestro •• 5567';
  
  // Quick Debug: logging to Debug sheet
  /*var debugSheet = SpreadsheetApp.getActive().getSheetByName("Debug");
  var debugSheetRow = Math.max(debugSheet.getLastRow(), 1) + 1;
  var debugSheetColumn = 1;
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")); SpreadsheetApp.flush();
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(subject); SpreadsheetApp.flush();
  debugSheet.getRange(debugSheetRow, debugSheetColumn++).setValue(message); SpreadsheetApp.flush();*/

  var tokenTypesWithCategories = {
    "Зачисление": availableToBudget,
    "Выдача": "↕️ Account Transfer",
    "Покупка": null,
    "Оплата": null,
    "Перевод": null
  };

  var tokenBalance = "Баланс=";
  var tokenCurrency = "₽";

  // Replace line breaks with spaces
  message = message.split("\n").join(" ");
  message = message.split("\r").join(" ");
  // Replace non-breaking spaces with regular spaces
  message = message.split(String.fromCharCode(160)).join(" ");
  // Remove duplicate spaces
  message = message.replace(/\s\s+/g, ' ');

  //var messageLowerCase = message.toLowerCase();

  // Filter our irrelevant messages
  var tokenFilters = [
    "СПАСИБО от Сбербанка",
    "ОТКАЗ",
    "Отказ",
    "Код для оплаты",
    "промокод",
    "Инициализация",
    "Встроенное средство защиты работает в фоновом режиме"
  ];
  var tokenFilterFound = false;
  tokenFilters.forEach(function(tokenFilter) {
    if (message.indexOf(tokenFilter) != -1) {
      tokenFilterFound = true;
      return; // this is return only from this small anonymous (unnamed) function
    }
  });
  if (tokenFilterFound) {
    return {error:"filtered"};
  }
  
  // Add info to "Memo" field
  var memo = subject + " | " + message.slice(0, message.indexOf(tokenBalance) - 1);
  var tokenExtraStart = message.indexOf(tokenCurrency, message.indexOf(tokenBalance)) + 1;
  if (tokenExtraStart != -1) {
    memo = memo + message.slice(tokenExtraStart);
  }
  
  // Determine category
  var category = "";
  var keyword = memo;
  for (var tokenType in tokenTypesWithCategories) {
    if (subject.indexOf(tokenType) != -1) {
      category = tokenTypesWithCategories[tokenType];
      if (category == null) {
        var category = categorize(keyword);
      }
      break;
    }
  }

  var amount = message.slice(0, message.indexOf(tokenCurrency) - 1).trim();
  // Erase spaces
  amount = amount.split(" ").join("");

  // Define outflow / inflow
  if (category == availableToBudget) {
    var outflow = "";
    var inflow = amount;
  } else {
    var outflow = amount;
    var inflow = "";
  }
  
  if (date == null) {
    // Take current date
    date = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy");
  }

  var account = "";
  if (message.indexOf("•• 5567") != -1) {
    account = "💰 Alexander’s Sberbank Maestro";
  } else if (message.indexOf("•• 1457") != -1) {
    account = "💰 Anna’s Sberbank MasterCard";
  }

  // TODO: implement parsing of balance amount in the message to perform check if the records are reconciled
  var balance = "0";
  var comment = isReconciled(account, Number(balance.split(",").join("."))) ? "Reconciled" : "Not reconciled";

  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:account, memo:memo, comment:comment, keyword:keyword};
  } else {
    return null;
  }
}

// Parsing of Sberbank outflow transaction email
// Though as Sberbank sends emails only for outflow transactions, parseSberbankSMS() is recommended to be used instead
function parseSberbankOutflowEmail(subject, message, date) {
  //var message = "Уважаемый(-ая), СЕМЁН СЕМЁНОВИЧ! Информируем Вас о том, что с Вашей карты Maestro **** 5427 выполнен платеж на карту **** 4212 на сумму 1,23 RUB.\n\n\nДля получения более подробной информации, а также для настройки параметров оповещений Вы можете воспользоваться системой Сбербанк Онл@йн. Это электронное сообщение содержит конфиденциальную информацию. Настоящим уведомляем Вас о том, что, если это сообщение не предназначено Вам, использование, копирование, распространение информации, содержащейся в настоящем сообщении, а также осуществление любых действий на основе этой информации, строго запрещено. Если Вы получили это сообщение по ошибке, пожалуйста, сообщите об этом и удалите это сообщение. Вы можете отказаться от получения уведомлений и подписаться на них вновь в настройках Сбербанк Онлайн. Пожалуйста, обратите внимание, что электронный адрес online@sberbank.ru, с которого отправлено данное сообщение, не предназначен для ответа. Пожалуйста, обращайтесь в отделения Сбербанка или по бесплатному телефону Контактного центра 900 (с мобильного по РФ) или +7 (495) 500-55-50, и мы с удовольствием ответим на все возникшие у Вас вопросы. Дата и время составления отчета: 04.11.2019 15:42:56 С уважением, Сбербанк России Адрес: 117997, г. Москва, ул. Вавилова д. 19. Телефон: +7 (495) 500-55-50}";
  var tokenMemo = "платеж ";
  var tokenAmount = "на сумму ";
  var tokenDate = "Дата и время составления отчета: ";
  var tokenExtra = "Для получения более подробной информации";

  // Replace line breaks with spaces
  message = message.split("\n").join(" ");
  message = message.split("\r").join(" ");
  // Replace non-breaking spaces with regular spaces
  message = message.split(String.fromCharCode(160)).join(" ");
  // Remove duplicate spaces
  message = message.replace(/\s\s+/g, ' ');

  // Find transaction info in the message and put it to "Memo"
  var memo = message.slice(message.indexOf(tokenMemo), message.indexOf(tokenExtra) - 1).trim();
  // Erase ending dot.
  memo = memo.slice(0, memo.length - 1);
  
  // Find date in the message
  if (message.indexOf(tokenDate) != -1) {
    var dateStart = message.indexOf(tokenDate) + tokenDate.length;
    date = message.substr(dateStart, 10).trim();
  } else if (date == null) {
    // Take current date
    date = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy");
  }

  // Find amount in the message
  var amountStart = message.indexOf(tokenAmount) + tokenAmount.length;
  var amountEnd = message.indexOf(tokenExtra, amountStart) - 1;
  var amount = message.slice(amountStart, amountEnd).trim();
  // Erase ending dot.
  amount = amount.slice(0, amount.length - 1);
  // Apply currency conversion if needed
  amount = parseAmountWithCurrency(amount, date);

  // Define outflow / inflow
  var outflow = amount;
  var inflow = "";
  
  var account = "";
  if (message.indexOf("**** 5567") != -1) {
    account = "💰 Alexander’s Sberbank Maestro";
  } else if (message.indexOf("**** 1457") != -1) {
    account = "💰 Anna’s Sberbank MasterCard";
  }

  // TODO: implement parsing of balance amount in the message to perform check if the records are reconciled
  var balance = "0";
  var comment = isReconciled(account, Number(balance.split(",").join("."))) ? "Reconciled" : "Not reconciled";

  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:account, memo:memo, comment:comment};
  } else {
    return null;
  }
}

// Parsing of Citibank Russia email
function parseCitibankEmail(subject, message, date) {
  // Quick Debug: examples of input variable
  //var message = "Dear Customer, The following transaction has been charged to credit card number **7001:\nAmount: 1,033.00 RUR\n\n\nPoint of sale: STARBUCKS         MO\nDate: 11/04/2019\nAvailable limit: 92,606.18 RUB The amount has been placed on hold in your credit card account. The actual debit will take place once a confirmation from the payment system has been received. If a transaction involves currency conversion, the actual charge to your account may differ from the amount placed on hold due to different exchange rates on the date of transaction and the date of actual debit. You can check the final transaction amount and the date of actual debit in Citibank Online or your monthly credit card statement. For more information on the card payments, conversion and debiting rules, please see our website at www.citibank.ru/cc/conversion/. If you didnвЂ™t authorize this transaction, please block your card and have it reissued. Details on how to do it are available here: https://www.citibank.ru/russia/pdf/pdf_instruction/Card_blocking_and_re-issue.pdf. If you disagree with the transaction, you can contact the bank to initiate a dispute. Details on how to do it are available here: https://www.citibank.ru/russia/pdf/pdf_instruction/Dispute.pdf. Learn more about Citibank Alerting Service on our website at www.citibank.ru. Sincerely, AO Citibank PLEASE DO NOT REPLY TO THIS MESSAGE. Please let us know of any changes in your contact details by signing on to CitibankВ® Online and choosing вЂњContact InformationвЂќ under вЂњMy ProfileвЂќ. You can set up Citibank Alerting Service preferences or opt out of receiving alerts by choosing вЂњCitibank Alerting ServiceвЂќ under вЂњProducts & ServicesвЂќ.";
  //var message = "\r\nDear Customer,\r\n\r\nYou have received a payment to your card:\r\nCard number **7001\r\nAmount: 5,000.00 RUB \r\nAvailable limit: 96,635.94 RUB\r\n \r\nAvailable cash limit: 96,635.94 RUB  \r\n\r\nIf you didn\u0432\u0402\u2122t authorize this transaction, please call CitiPhone immediately on +7(495)775-75-75 in Moscow, +7(812)336-75-75 in St. Petersburg or 8(800)700-38-38 elsewhere in Russia.\r\n\r\nInformation on how to dispute a charge to your account is available by visiting the \u0432\u0402\u045aFAQ\u0432\u0402\u045c page under \u0432\u0402\u045aContact Us\u0432\u0402\u045c at www.citibank.ru, or by clicking https://www.citibank.ru/russia/pdf/dispute_leaflet_rus.pdf  (how to dispute a charge to your debit or credit card) and ?         https://www.citibank.ru/russia/info/rus/pdf/disput_form_01-2017.pdf (Transaction Dispute Form).\r\n\r\nLearn more about Citibank Alerting Service on our website at www.citibank.ru.\r\n\r\nSincerely,\r\nAO Citibank\r\n\r\nPLEASE DO NOT REPLY TO THIS MESSAGE.\r\n  \r\nPlease let us know of any changes in your contact details by signing on to Citibank\u0412\u00ae Online and choosing \u0432\u0402\u045aContact Information\u0432\u0402\u045c under \u0432\u0402\u045aMy Profile\u0432\u0402\u045c. \r\nYou can set up Citibank Alerting Service preferences or opt out of receiving alerts by choosing \u0432\u0402\u045aCitibank Alerting Service\u0432\u0402\u045c under \u0432\u0402\u045aProducts & Services\u0432\u0402\u045c.";
  //var message = "\r\nDear Customer,\r\n\r\nYou have received a payment to your credit card:\r\nCard number **7001\r\nAmount: 5,000.00 RUB \r\nAvailable limit: 96,635.94 RUB\r\n \r\nAvailable cash limit: 96,635.94 RUB  \r\n\r\nIf you didn\u0432\u0402\u2122t authorize this transaction, please call CitiPhone immediately on +7(495)775-75-75 in Moscow, +7(812)336-75-75 in St. Petersburg or 8(800)700-38-38 elsewhere in Russia.\r\n\r\nInformation on how to dispute a charge to your account is available by visiting the \u0432\u0402\u045aFAQ\u0432\u0402\u045c page under \u0432\u0402\u045aContact Us\u0432\u0402\u045c at www.citibank.ru, or by clicking https://www.citibank.ru/russia/pdf/dispute_leaflet_rus.pdf  (how to dispute a charge to your debit or credit card) and ?         https://www.citibank.ru/russia/info/rus/pdf/disput_form_01-2017.pdf (Transaction Dispute Form).\r\n\r\nLearn more about Citibank Alerting Service on our website at www.citibank.ru.\r\n\r\nSincerely,\r\nAO Citibank\r\n\r\nPLEASE DO NOT REPLY TO THIS MESSAGE.\r\n  \r\nPlease let us know of any changes in your contact details by signing on to Citibank\u0412\u00ae Online and choosing \u0432\u0402\u045aContact Information\u0432\u0402\u045c under \u0432\u0402\u045aMy Profile\u0432\u0402\u045c. \r\nYou can set up Citibank Alerting Service preferences or opt out of receiving alerts by choosing \u0432\u0402\u045aCitibank Alerting Service\u0432\u0402\u045c under \u0432\u0402\u045aProducts & Services\u0432\u0402\u045c.";
  //var message = "Dear Customer,\r\n\r\nTransaction on supplementary card **7001:\r\nAmount: 229.00 RUB \r\nPoint of sale: STARBUCKS         MO\r\nDate: 12/03/2019\r\nAvailable limit: 92,771.00 RUB \r\n\r\nThe amount has been placed on hold in your credit card account. The actual debit will take place once a confirmation from the payment system has been received. \r\nIf a transaction involves currency conversion, the actual charge to your account may differ from the amount placed on hold due to different exchange rates on the date of transaction and the date of actual debit. \r\nYou can check the final transaction amount and the date of actual debit in Citibank Online or your monthly credit card statement. \r\nFor more information on the card payments, conversion and debiting rules, please see our website at www.citibank.ru/cc/conversion/.\r\n\r\nIf you didn\u0432\u0402\u2122t authorize this transaction, please call CitiPhone immediately on +7(495)775-75-75 in Moscow, +7(812)336-75-75 in St. Petersburg or 8(800)700-38-38 elsewhere in Russia.\r\n\r\nInformation on how to dispute a charge to your account is available by visiting the \u0432\u0402\u045aFAQ\u0432\u0402\u045c page under \u0432\u0402\u045aContact Us\u0432\u0402\u045c at www.citibank.ru, or by clicking https://www.citibank.ru/russia/pdf/dispute_leaflet_rus.pdf  (how to dispute a charge to your debit or credit card) and ?         https://www.citibank.ru/russia/info/rus/pdf/disput_form_01-2017.pdf (Transaction Dispute Form).\r\n\r\nLearn more about Citibank Alerting Service on our website at www.citibank.ru.\r\n\r\nSincerely,\r\nAO Citibank\r\n\r\nPLEASE DO NOT REPLY TO THIS MESSAGE.\r\n  \r\nPlease let us know of any changes in your contact details by signing on to Citibank\u0412\u00ae Online and choosing \u0432\u0402\u045aContact Information\u0432\u0402\u045c under \u0432\u0402\u045aMy Profile\u0432\u0402\u045c. \r\nYou can set up Citibank Alerting Service preferences or opt out of receiving alerts by choosing \u0432\u0402\u045aCitibank Alerting Service\u0432\u0402\u045c under \u0432\u0402\u045aProducts & Services\u0432\u0402\u045c.";
  //var message = "Dear Customer, Amount: 1,005.00 RUB has been charged back to your card **7001, Transaction: UBER.             MO. Available balance: 94,227.44 RUB. If you didn’t authorize this transaction, please call CitiPhone immediately on +7(495)775-75-75 in Moscow, +7(812)336-75-75 in St. Petersburg or 8(800)700-38-38 elsewhere in Russia. Information on how to dispute a charge to your account is available by visiting the “FAQ” page under “Contact Us” at www.citibank.ru, or by clicking https://www.citibank.ru/russia/pdf/dispute_leaflet_rus.pdf  (how to dispute a charge to your debit or credit card) and ?         https://www.citibank.ru/russia/info/rus/pdf/disput_form_01-2017.pdf (Transaction Dispute Form). Learn more about Citibank Alerting Service on our website at www.citibank.ru. Sincerely, AO Citibank PLEASE DO NOT REPLY TO THIS MESSAGE. Please let us know of any changes in your contact details by signing on to Citibank® Online and choosing “Contact Information” under “My Profile”. You can set up Citibank Alerting Service preferences or opt out of receiving alerts by choosing “Citibank Alerting Service” under “Products & Services”.";
  var tokenAmount = "Amount: ";
  var tokenPointOfSale = "Point of sale: ";
  var tokenDate = "Date: ";
  var tokenBalance = "Available limit: ";
  if (message.indexOf(tokenBalance) == -1) {
    tokenBalance = "Available balance: ";
  }
  var tokenChargedBack = "has been charged back";
  /*var tokenAmount = "Сумма: ";
  var tokenPointOfSale = "Торговая точка: ";
  var tokenDate = "Дата операции: ";
  var tokenBalance = "Доступный лимит: ";
  if (message.indexOf(tokenBalance) == -1) {
    tokenBalance = "Доступный баланс: ";
  }*/

  // Find transaction info in the message and put it to "Memo"
  var tokenDateStart = message.indexOf(tokenDate);
  if (tokenDateStart != -1) {
    var memo = message.slice(message.indexOf(tokenAmount), tokenDateStart - 1).replace("\n", " | ").trim();
  } else {
    var memo = message.slice(message.indexOf(tokenAmount), message.indexOf(tokenBalance) - 1).replace("\n", " | ").trim();
  }
  memo = memo.split("\n").join(" ");
  memo = memo.split("\r").join(" ");
  // Remove duplicate spaces
  memo = memo.replace(/\s\s+/g, ' ');
  
  // Find date in the message
  if (tokenDateStart != -1) {
    var dateStart = message.indexOf(tokenDate) + tokenDate.length;
    date = message.substr(dateStart, 10).trim();
    // Convert date from MM/DD/YYYY to DD.MM.YYYY
    var dateSplit = date.split("/");
    date = dateSplit[1] + "." + dateSplit[0] + "." + dateSplit[2];
  } else if (date == null) {
    // Take current date
    date = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy");
  }

  // Find amount in the message
  var amountStart = message.indexOf(tokenAmount) + tokenAmount.length;
  if (message.indexOf(tokenChargedBack, amountStart) != -1) {
    var amountEnd = message.indexOf(tokenChargedBack, amountStart) - 1;
  } else if (message.indexOf(tokenPointOfSale, amountStart) != -1) {
    var amountEnd = message.indexOf(tokenPointOfSale, amountStart) - 1;
  } else if (message.indexOf(tokenBalance, amountStart) != -1) {
    var amountEnd = message.indexOf(tokenBalance, amountStart) - 1;
  } else {
    var amountEnd = message.indexOf(tokenDate, amountStart) - 1;
  }
  var amount = message.slice(amountStart, amountEnd).trim();
  // Convert from X,XXX,XXX.XX to XXXXXXX.XX
  amount = amount.split(",").join("");
  // Convert from XXXXXXX.XX to XXXXXXX,XX
  amount = amount.replace(".", ",");
  // Apply currency conversion if needed
  amount = parseAmountWithCurrency(amount, date);

  // Define outflow, inflow and category
  var outflow = "";
  var inflow = "";
  var category = "";
  if ((message.indexOf("transaction has been charged") != -1) || (message.indexOf("Transaction on ") != -1)) { // "произведено списание"
    outflow = amount;
    var pointOfSale = message.slice(message.indexOf(tokenPointOfSale) + tokenPointOfSale.length, message.indexOf(tokenDate)).trim();
    var separator = "  ";
    if (pointOfSale.indexOf(separator) != -1) {
      pointOfSale = pointOfSale.slice(0, pointOfSale.indexOf(separator));
    }
    var keyword = pointOfSale;
    category = categorize(keyword);
  } else if ((message.indexOf("You have received a payment") != -1) || (message.indexOf("charged back") != -1)) {
    inflow = amount;
    category = availableToBudget;
  }
  
  var account = "💳 Alexander’s Citibank MasterCard";
  
  if (message.indexOf("supplementary card") != -1) {
    memo = memo + " | Transaction on supplementary card";
  }
  
  // Find balance in the message
  var balanceStart = message.indexOf(tokenBalance) + tokenBalance.length;
  var balanceEnd = message.indexOf(" ", balanceStart) + 4; // 3-character currency code + 1
  var balance = message.slice(balanceStart, balanceEnd).trim();
  // Convert from X,XXX,XXX.XX to XXXXXXX.XX
  balance = balance.split(",").join("");
  // Convert from XXXXXXX.XX to XXXXXXX,XX
  balance = balance.replace(".", ",");
  // Apply currency conversion if needed
  balance = parseAmountWithCurrency(balance, date);

  var comment = "";
  
  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:account, memo:memo, comment:comment, keyword:keyword};
  } else {
    return null;
  }
}

// Parsing of Citibank Russia SMS
// Though as Citibank sends more information in e-mails than in SMS, parseCitibankEmail() is recommended to be used instead
function parseCitibankSMS(message, date) {
  // Quick Debug: examples of input variable
  //var message = 'Card *7001; 44.76USD;WALGREENS #;06/13; Available 89933.83RUB';
  //var message = 'Card *7001; 18.00GBP;TFL CYCLE H;06/11; Available 97093.16RUB';
  //var message = 'Card *7001; 608.00RUB;UBER       ;08/08; Available 2229.97RUB';
  //var message = 'Card *7001 Deposit 5000.00 RUB Available 85656.75 RUB';
  //var message = 'Payment to card account **7001; Amount: 1,692 RUR; Date: 06/21/18';
  //var message = 'Transaction on supplementary credit card **7001 Amount: 4,753.72 RUB  Point of sale: MOSOBLEIRC             PR Date: 06/10/18 Available limit: 89,197.40 RUB';
  // Сообщения по старому формату 2017:
  //var message = 'Transaction on credit card **7004 Amount: 51.00 RUB  Point of sale: ROSTELEKOM             MO Date: 09/06/17 Available limit: 86,478.45 RUB';
  //var message = 'Payment to credit card **7004 Amount: 5,000.00 RUB Available limit: 82,478.45 RUB, Available cash limit: 82,478.45 RUB';
  // Сообщения по старому формату 2016:
  //var message = 'Operatsiya po kreditnoy karte **7009 Summa: 2,981.00 RUB Torgovaya tochka: AUCHAN MO Data: 06/01/16 Dostupniy limit: 85,571.11 RUB';
  // Сообщения по старому формату 2015:
  //var message = 'Spisanie po kreditnoy karte **7005 Summa: 470.00 RUB Torgovaya tochka: PAYMENT Data: 30/05/15 Dostupniy limit: 89,285.69 RUB';
  // Сообщения по старому формату 2014:
  //var message = 'Payment to credit card **7008\nAmount: 5,000.00 RUB\nAvailable limit 85,569.39 RUB\nAvailable cash limit 85,569.39 RUB';
  //var message = 'Debit from credit card **7008\nAmount: EUR  12.00\nPoint of sale: VIVATICKET.IT          SA\nDate: 08/01/11\nAvailable limit: RUB 89,174.12';
  //var message = 'Debit from credit card **7008\nAmount: RUB  300.93\nPoint of sale: PAYPAL *WIKIMEDIAFO    41\nDate: 09/09/11\nAvailable limit: RUB 87,262.46';
  //var message = 'Debit from credit card **7008\nAmount: RUR 600.00\nPoint of sale: SI PAYMENT 9101221221\nDate: 02/15/14\nAvailable limit: RUB 89,969.39';
  // Сообщения по старому формату 2011:
  //var message = 'Debit: 0.99 USD ; credit card: 7007; transaction: Master AM; date: 25/07/2010;  available balance: 85,899.80.';
  //var message = 'Debit: 150.00 RUR ; credit card: 7007; transaction: PAYMENT 9101221221       ; date: 26/07/2010;  available balance: 85,749.80.';
  //var message = 'Credit: 2,000.00 RUR; transaction: CC payment; credit card: 7007  date: 14/07/2010.';
  //var message = 'Spisanie: 1.00 USD ; kreditnaya karta: 7007; operacija: PAYPAL                 40; data: 2009/12/30;  dostupnii limit: 86,669.00.';
  
  var tokenAmount = "Amount: ";
  if (message.indexOf(tokenAmount) == -1) tokenAmount = "Summa: ";
  if (message.indexOf(tokenAmount) == -1) tokenAmount = "Debit: ";
  if (message.indexOf(tokenAmount) == -1) tokenAmount = "Credit: ";
  if (message.indexOf(tokenAmount) == -1) tokenAmount = "Spisanie: ";
  var tokenPointOfSale = "Point of sale: ";
  if (message.indexOf(tokenPointOfSale) == -1) tokenPointOfSale = "Torgovaya tochka: ";
  if (message.indexOf(tokenPointOfSale) == -1) tokenPointOfSale = "transaction: ";
  if (message.indexOf(tokenPointOfSale) == -1) tokenPointOfSale = "operacija: ";
  var tokenDate = "Date: ";
  if (message.indexOf(tokenDate) == -1) tokenDate = "Data: ";
  if (message.indexOf(tokenDate) == -1) tokenDate = "date: ";
  if (message.indexOf(tokenDate) == -1) tokenDate = "data: ";
  var tokenBalance = "Available limit: ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "Available limit ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "Available balance: ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "Available balance ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "Dostupniy limit: ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "available balance: ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "dostupnii limit: ";
  if (message.indexOf(tokenBalance) == -1) tokenBalance = "Available ";
  var tokenDeposit = "Deposit ";
  var tokenChargedBack = "has been charged back";
  var tokenSeparator = ";";

  // Filter our irrelevant messages
  // Security Alert: failed Sign On attempt to Citibank Online/Citi Mobile 2018.06.15 11:01 Moscow time
  // Security Alert: successful Sign On to Citibank Online 26/10/2013 14:20 Moscow time
  // Successful logon to Citibank Online 22/09/2012 12:26
  // Uspeshniy vhod v Citibank Online/Citi Mobile 2014.11.17 12:51 po moskovskomu vremeni. Otpravleno v celjah bezopasnosti
  // Neuspeshnaya popytka vhoda v Citibank Online/Citi Mobile 2017.02.15 16:47 po moskovskomu vremeni. Otpravleno v celjah bezopasnosti
  // Dear Customer, Samsung Pay has been successfully activated for your card **7001. Your unique digital card number is XX4565; you can also find it in Samsung Pay. AO Citibank
  // Dear customer, you have set up a new PIN for your credit card *7001. If you did not do this, contact Citibank immediately! To have your PIN recorded on the card chip, use the card at any ATM, e.g., make a balance enquiry. AO Citibank.
  // Dear customer, this message is to confirm that your Citibank Online password has been reset.
  // Dear Customer! Your address has updated in our system based on your request. If you have not requested this, please call Citiphone.
  var tokenFilters = [
    "failed Sign On attempt",
    "successful Sign On"
  ];

  // Find transaction info in the message and put it to "Memo"
  var tokenDateStart = message.indexOf(tokenDate);
  var memo = message;
  memo = memo.split("\n").join(" ");
  memo = memo.split("\r").join(" ");
  // Remove duplicate spaces
  memo = memo.replace(/\s\s+/g, ' ');
  
  // Find amount in the message
  if (message.indexOf(tokenAmount) != -1) {
    var amountStart = message.indexOf(tokenAmount) + tokenAmount.length;
    if (message.indexOf(tokenChargedBack, amountStart) != -1) {
      var amountEnd = message.indexOf(tokenChargedBack, amountStart) - 1;
    } else if (message.indexOf(tokenPointOfSale, amountStart) != -1) {
      var amountEnd = message.indexOf(tokenPointOfSale, amountStart) - 1;
    } else if (message.indexOf(tokenBalance, amountStart) != -1) {
      var amountEnd = message.indexOf(tokenBalance, amountStart) - 1;
    } else {
      var amountEnd = message.indexOf(tokenDate, amountStart) - 1;
    }
  } else if (message.indexOf(tokenDeposit) != -1) {
    var amountStart = message.indexOf(tokenDeposit) + tokenDeposit.length;
    var amountEnd = message.indexOf(tokenBalance, amountStart);
  } else {
    var amountStart = message.indexOf(tokenSeparator) + tokenSeparator.length;
    var amountEnd = message.indexOf(tokenSeparator, amountStart);
  }
  var amount = message.slice(amountStart, amountEnd).trim();
  // Remove duplicate spaces
  amount = amount.replace(/\s\s+/g, ' ');
  // Erase separator at the end if there is one
  if (amount.substr(-1) == tokenSeparator) {
    amount = amount.slice(0, amount.length - 1).trim();
  }
  // Swap currency and amount in case of reverse notation (RUR 500.00)
  if (isNaN(amount.slice(0, 1))) {
    var amountSplit = amount.split(" ");
    amount = amountSplit[1] + " " + amountSplit[0];
  }
  // Convert from X,XXX,XXX.XX to XXXXXXX.XX
  amount = amount.split(",").join("");
  // Convert from XXXXXXX.XX to XXXXXXX,XX
  amount = amount.replace(".", ",");
  // Apply currency conversion if needed
  amount = parseAmountWithCurrency(amount, date);

  // Define outflow, inflow and category
  var outflow = "";
  var inflow = "";
  var category = "";
  if ((message.indexOf("Payment to ") != -1) || (message.indexOf("Deposit ") != -1)) {
    inflow = amount;
    category = availableToBudget;
  } else {
    outflow = amount;
  }
  
  var account = "💳 Alexander’s Citibank MasterCard";
  
  var comment = "";
  var keyword = "";
  
  //TODO: add error checking
  if (true) {
    return {date:date, outflow:outflow, inflow:inflow, category:category, account:account, memo:memo, comment:comment, keyword:keyword};
  } else {
    return null;
  }
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Transaction automatic categorization functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function categorize(keyword) {
  if (keyword.trim() == "") {
    return "";
  }
  var catSheet = SpreadsheetApp.getActive().getSheetByName("Keywords & Categories");
  var firstRow = 2;
  var lastRow = Math.max(catSheet.getLastRow(), 1);
  var numRows = lastRow - firstRow + 1;
  
  keyword = keyword.toLowerCase();
  
  var data = catSheet.getRange(firstRow, 1, lastRow, 2).getValues(); // create an array of data from columns A and B

  // Try to find exact match first
  for (var nn = 0; nn < numRows; ++nn) {
    if (data[nn][0] == keyword) { // if a match in column A is found, break the loop
      break;
    };
  }
  var mostPopularCategory = data[nn][1];
  if ((nn < numRows) && (mostPopularCategory != "")) {
    return mostPopularCategory;
  };
  
  // If exact match was not found, try to find a partial match
  for (var nn = 0; nn < numRows; ++nn) {
    if ((data[nn][0].indexOf(keyword) != -1) || (keyword.indexOf(data[nn][0]) != -1)) {
      break;
    };
  }
  var mostPopularCategory = data[nn][1];
  if ((nn < numRows) && (mostPopularCategory != "")) {
    return mostPopularCategory;
  };
  
  return "";
}

function batchCategorize() {
  var spreadsheet = SpreadsheetApp.getActive();
  //var sheet = spreadsheet.getActiveSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName("Transactions");
  var lastColumn = sheet.getMaxColumns();
  var firstRow = 8;
  var lastRow = Math.max(sheet.getLastRow(), 1) + 1;

  var keywordArrayColumn = 9; // Note: array column numbers start from 0
  
  var categorySheetColumn = 4 + 1; // Note: sheet column numbers start from 1
  var keywordSheetColumn = keywordArrayColumn + 1;
  
  var keywords = sheet.getRange(1, keywordSheetColumn, lastRow, keywordSheetColumn).getValues();

  var category = "";
  
  for (var currentRow = firstRow; currentRow < lastRow; currentRow++) {
    if (!sheet.isRowHiddenByFilter(currentRow) && !sheet.isRowHiddenByUser(currentRow)) {
      var keyword = keywords[currentRow - 1][0];
      category = categorize(keyword);
      sheet.getRange(currentRow, categorySheetColumn).setValue(category);
      sheet.getRange(currentRow, categorySheetColumn - 1).setValue(keyword);
    }
  }
  SpreadsheetApp.flush();
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Transaction automatic reconcilation functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function isReconciled(account, balance) {
  var calcSheet = SpreadsheetApp.getActive().getSheetByName("Calculations");
  var delta = 300; // in your main currency, maximum allowed difference between balance in bank message and current balance in Aspire Budget (useful  for transactions in foreign currency not to trigger "Not reconciled")
  var firstRow = 52;
  var lastRow = Math.max(calcSheet.getLastRow(), 1);
  var numRows = lastRow - firstRow + 1;
  
  var data = calcSheet.getRange(firstRow, 1, lastRow, 2).getValues(); // create an array of data from columns A and B
  
  for (var nn = 0; nn < numRows; ++nn) {
    if (data[nn][0] == account) { break }; // if a match in column A is found, break the loop
  }
  if (nn >= numRows) { return false };
  return ((balance - data[nn][1]) < delta);
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Multiple currency support functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function parseAmountWithCurrency(amount, date) {
  if (isNaN(amount.substr(-1))) {
    var RUBSymbols = ["₽", "р", "р.", "RUB", "RUR", "руб", "руб.", "рубль", "рубля", "рублей"];
    var RUBSymbolFound = "";
    RUBSymbols.forEach(function(RUBSymbol) {
      if (amount.indexOf(RUBSymbol) != -1) {
        RUBSymbolFound = RUBSymbol;
        return; // this is return only from this small anonymous (unnamed) function
      }
    });
    
    if (RUBSymbolFound) {
      var currency = "RUB";
      amount = amount.substring(0, amount.length - RUBSymbolFound.length);
    } else if (amount.indexOf(" ") != -1) {
      var currency = amount.split(" ")[1];
      amount = amount.split(" ")[0];
    } else {
      var currency = amount.substr(-3);
      amount = amount.substring(0, amount.length - 3);
    }
    if (currency != "RUB") {
      amount = '=' + amount + '*INDEX(GOOGLEFINANCE("CURRENCY:' + currency + 'RUB"; "close"; "' + date + '"; 3); 2; 2)'; // If this formula results in #ERROR, then you have another locale, so you might need to change ; to , and/or " to '
    }
  }
  return amount;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * HTML processing functionality
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function getElementById(element, idToFind) {  
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();
    if( elt != null) {
      var id = elt.getAttribute('id');
      if( id != null && id.getValue() == idToFind) return elt;    
    }
  }
}

function getElementsByClassName(element, classToFind) {  
  var data = [];
  var descendants = element.getDescendants();
  descendants.push(element);  
  for(i in descendants) {
    var elt = descendants[i].asElement();
    if(elt != null) {
      var classes = elt.getAttribute('class');
      if(classes != null) {
        classes = classes.getValue();
        if(classes == classToFind) data.push(elt);
        else {
          classes = classes.split(' ');
          for(j in classes) {
            if(classes[j] == classToFind) {
              data.push(elt);
              break;
            }
          }
        }
      }
    }
  }
  return data;
}

function getElementsByTagName(element, tagName) {  
  var data = [];
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();     
    if( elt != null && elt.getName() == tagName) data.push(elt);      
  }
  return data;
}

function getElementByText(element, textToFind) {  
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();
    if( elt != null) {
      var text = elt.getText();
      if( text != null && text.getValue() == textToFind) return elt;    
    }
  }
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
* Macros for adding second row for "↕️ Account Transfer" transaction
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

function AddSecondRowForAccountTransfer() {
  var spreadsheet = SpreadsheetApp.getActive();
  //var sheet = spreadsheet.getActiveSheet();
  var sheet = SpreadsheetApp.getActive().getSheetByName("Transactions");
  var activeRanges = sheet.getActiveRangeList().getRanges();
  var lastColumn = sheet.getMaxColumns();

  var outflowArrayColumn = 2; // Note: array column numbers start from 0
  var inflowArrayColumn = 3;
  var accountArrayColumn = 5;
  
  var categorySheetColumn = 4 + 1; // Note: sheet column numbers start from 1
  var timestampSheetColumn = 0 + 1;

  for (var i = activeRanges.length - 1; i >= 0; i--) {
    var rangeFirstRow = activeRanges[i].getRow();
    var rangeLastRow = rangeFirstRow + activeRanges[i].getNumRows() - 1;
    
    for (var currentRow = rangeLastRow; currentRow >= rangeFirstRow; currentRow--) {
      if (!sheet.isRowHiddenByFilter(currentRow) && !sheet.isRowHiddenByUser(currentRow)) {
        sheet.getRange(currentRow, categorySheetColumn).setValue("↕️ Account Transfer");
        // Save current timestamp as a unique ID that can be used as a value that binds both lines
        sheet.getRange(currentRow, timestampSheetColumn).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"));
        
        var rangeData = getValuesAndFormulas(sheet.getRange(currentRow, 1, 1, lastColumn));
        var outflow = rangeData[0][outflowArrayColumn];
        var inflow = rangeData[0][inflowArrayColumn];
        rangeData[0][outflowArrayColumn] = inflow;
        rangeData[0][inflowArrayColumn] = outflow;
        // Replace "💰 John’s First Bank Account" with "💰 John’s Cash"
        var account = rangeData[0][accountArrayColumn];
        rangeData[0][accountArrayColumn] = account.slice(0, account.indexOf(" ", 3)) + " Cash";
        spreadsheet.toast("Processing row " + currentRow, Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"), 3);
        
        sheet.insertRowAfter(currentRow);
        sheet.getRange(currentRow + 1, 1, 1, lastColumn).setValues(rangeData);
      }
    }
  }
  SpreadsheetApp.flush();
}

function getValuesAndFormulas(range) {
  var rangeData = range.getValues();
  var rangeFormulas = range.getFormulas();

  for (lin in rangeFormulas)
    for (col in rangeFormulas[lin])
      if (rangeFormulas[lin][col] != "")
        rangeData[lin][col] = rangeFormulas[lin][col];

  return rangeData;
}
