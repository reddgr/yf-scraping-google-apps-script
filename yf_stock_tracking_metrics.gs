// RETURNS CURRENT STOCK PRICE:
function Yfinance(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}?p=${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const price = contentText.match(/<fin-streamer(?:.*?)active="">(\d+[,]?[\d\.]+?)<\/fin-streamer>/);
  return price[1];
}
// TEST:
console.log("AAPL current price: " + Yfinance("AAPL"))

// RETURNS MARKET CAP:
function YFmc(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}?p=${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const marketcap = contentText.match(/data-test="MARKET_CAP-value">(.*?)<\/td>/);
  return marketcap[1];
}
//console.log(YFmc("AAPL"))

function YFbeta(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}?p=${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const marketcap = contentText.match(/data-test="BETA_5Y-value">(.*?)<\/td>/);
  return marketcap[1];
}

// TEST:
console.log("GOOG current beta: " + YFbeta("GOOG"))

function YFper(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}?p=${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const peRatio = contentText.match(/data-test="PE_RATIO-value">(.*?)<\/td>/);
  return peRatio[1];
}
console.log("TCTZF PE ratio: " + YFper("TCTZF"))

function YFgrowth(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}/analysis`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const growth5yr = contentText.match(/Past 5 Years \(per annum\)<\/span><\/td><td class="Ta\(end\) Py\(10px\)">(-?\d+\.\d+%)/);
  return growth5yr ? growth5yr[1] : "NA";
}
console.log("IBM 5 yr growth: " + YFgrowth("IBM"))

function getThirdFinColValue(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}/balance-sheet`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();  
  // Match all instances of the pattern, then select the third one
  const pattern = /<div class="[^"]*" data-test="fin-col"><span>([\d,]+)<\/span><\/div>/g;
  let match;
  let count = 0;
  let thirdValue = "Value not found";
  
  // Loop through all matches to find the third occurrence
  while ((match = pattern.exec(contentText)) !== null) {
    count++;
    if (count === 3) {
      thirdValue = match[1]; // This is the third match's captured group
      break;
    }
  }  
  return thirdValue;
}
console.log("IBM test: " + getThirdFinColValue("IBM"))


function extractGrossMargin(ticker) {
  const url = `https://example.com/symbol/${ticker}`; // Adjust with the actual URL
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = response.getContentText();

  // Updated regex to be more flexible with spaces and potential HTML encoding
  const regexPattern = /Gross Margin is\s*([\d.,]+%)/;
  const match = contentText.match(regexPattern);

  // Check if a match was found and return the value, otherwise indicate not found
  return match ? match[1] : "Gross Margin percentage not found";
}
// Example usage
const grossMargin = extractGrossMargin("NVDA");
console.log("NVDA Gross Margin:", grossMargin);

// DUPLICATE A SHEET WITH VALUES TO KEEP A MANUAL BACKUP OF YAHOO FINANCE METRICS
function duplicateSheet() {
 // Get the source sheet to duplicate
 var sourceSheetName = "DGR12"; // Replace with the actual name of your sheet
 var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName);
 // Get today's date in a simple format (YYYYMMDD)
 var today = Utilities.formatDate(new Date(), "GMT", "yyyyMMdd");

 // Create a new sheet name with source sheet name and today's date appended
 var newSheetName = sourceSheetName + "-" + today;
 
 // Create a new sheet with a descriptive name
 //var newSheetName = "Copy of " + sourceSheetName;
 var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newSheetName);

 // Get the last row and column of the source sheet to define the range
 var sourceLastRow = sourceSheet.getLastRow();
 var sourceLastColumn = sourceSheet.getLastColumn();
 // Copy the source sheet's values and formatting to the new sheet, excluding formulas
 sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .copyTo(newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn),
             {contentsOnly: true, format: true});
 // Apply additional formatting for robustness
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setNumberFormats(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getNumberFormats());
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setFontWeights(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getFontWeights());
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setTextStyles(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getTextStyles());
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setBackgroundColors(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getBackgrounds());
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setHorizontalAlignments(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getHorizontalAlignments());
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setVerticalAlignments(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getVerticalAlignments());
 newSheet.getRange(1, 1, sourceLastRow, sourceLastColumn)
     .setWraps(sourceSheet.getRange(1, 1, sourceLastRow, sourceLastColumn).getWraps());
}
