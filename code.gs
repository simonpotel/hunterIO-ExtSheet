// Function to save the API key of Hunter.IO (has to be sent manually by the user)
function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
}
