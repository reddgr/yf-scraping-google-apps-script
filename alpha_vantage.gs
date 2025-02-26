function getLastClosePrice(ticker) {
  API_KEY = PropertiesService.getScriptProperties().getProperty("av_api_key")
  var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + ticker + "&apikey=" + API_KEY + "&outputsize=full");
  var json = response.getContentText();
  var data = JSON.parse(json);

  var timeSeries = data["Time Series (Daily)"];
  if (!timeSeries) {
    return "Time Series data not found";
  }

  var lastDate = Object.keys(timeSeries)[0]; // Get the most recent date
  var lastClose = timeSeries[lastDate]["4. close"]; // Extract the closing price

  return lastClose;
}

// TEST:
console.log(`AV - AAPL last close price: ${getLastClosePrice("AAPL")}`);

function getPERatio(ticker) {
  API_KEY = PropertiesService.getScriptProperties().getProperty("av_api_key");
  var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=OVERVIEW&symbol=" + ticker + "&apikey=" + API_KEY);
  var json = response.getContentText();
  var data = JSON.parse(json);

  return data["PERatio"] || "PERatio not found";
}

// TEST:
console.log(`AAPL PE Ratio: ${getPERatio("AAPL")}`);

function getOverview(ticker) {
  API_KEY = PropertiesService.getScriptProperties().getProperty("av_api_key")
  var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=OVERVIEW&symbol=" + ticker + "&apikey=" + API_KEY);
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  // Convert JSON object to a readable string format
  return JSON.stringify(data, null, 2);
}

// TEST:
console.log("AAPL overview: " + getOverview("AAPL"));