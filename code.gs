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
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Set Hunter.io API Key');
}

// Function to save the API key of Hunter.IO (has to be sent manually by the user)
function saveApiKey() {
  var apiKey = document.getElementById('apiKey').value;
  google.script.run
    .withSuccessHandler(function() {
      alert('API key saved successfully!');
    })
    .saveApiKey(apiKey);
  google.script.host.close();
}