function getJotformData() {
  var url = 'https://api.jotform.com/form/240147357265053/submissions?apiKey=619f0e30886ad49d75425eb7fa7f9d6e';
  var options = {
    'method' : 'get', 
    'contentType': 'application/json'
  };

  var response = UrlFetchApp.fetch(url, options);
  
  // Parse the JSON reply
  var json = response.getContentText();
  var data = JSON.parse(json);

  Logger.log(data);
}