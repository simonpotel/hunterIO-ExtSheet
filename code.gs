// Function to open the GUI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hunter.io')
    .addItem('Set API Key', 'showApiKeyDialog')
    .addToUi();
}

// Func of the Item in GUI 'Set Api Key', thats open a new dialog to save api key.
function showApiKeyDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('ApiKeyDialog')
    .setWidth(400) // size of the dialog window 
    .setHeight(300); // size of the dialog window
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Set Hunter.io API Key');
}

// Function to save the API key of Hunter.IO (has to be sent manually by the user)
function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
}

// Function to get the saved API key
function getApiKey() {
  return PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY') || '';
}

// Main function to find email addresses using Hunter.io API
function findEmail(firstName, lastName, companyDomain) {
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  if (!apiKey) {
    return 'API key not set.';
  }
  
  var url = `https://api.hunter.io/v2/email-finder?domain=${companyDomain}&first_name=${firstName}&last_name=${lastName}&api_key=${apiKey}`;
  
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); // request to the api  
    var data = JSON.parse(response.getContentText()); // content of the api request 
    
    if (response.getResponseCode() !== 200) {
      return `Error: ${data.errors[0].details}`; // manage error from the API 
    }
    
    if (data.data && data.data.email) {
      return data.data.email; 
    } else {
      return 'Email not found';
    }
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

// name function callback
function FindEmail(firstName, lastName, companyDomain) {
  return findEmail(firstName, lastName, companyDomain);
}

// Function to find emails using HunterIO api by process a range between CASES
function findEmailsInRange(range) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange(range).getValues();
  for (var i = 0; i < data.length; i++) {
    var firstName = data[i][0];
    var lastName = data[i][1];
    var companyDomain = data[i][2];
    var email = findEmail(firstName, lastName, companyDomain);
    sheet.getRange(i + 1, 4).setValue(email); // write result in the case thats writed the func 
  }
}