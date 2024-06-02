function saveApiKey() {
    var apiKey = document.getElementById('apiKey').value;
    google.script.run.saveApiKey(apiKey);
    google.script.host.close();
  }
  