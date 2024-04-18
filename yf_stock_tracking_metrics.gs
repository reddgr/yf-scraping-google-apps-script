// RETURNS CURRENT STOCK PRICE:
function getPrice(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const regex = /<fin-streamer[^>]*?class="livePrice[^>]*?data-value="([\d\.]+)"[^>]*?active>/;
  const priceMatch = contentText.match(regex);
  if (priceMatch && priceMatch[1]) {
    return priceMatch[1];  // Return the captured price
  }
  return 'NA';  // Return a default message if no price is captured
}
// TEST:
console.log("IBM current price: " + getPrice("IBM"))

// RETURNS CURRENT MARKET CAP:
function getMarketCap(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  // Regex to extract market cap from the data-value attribute of the <fin-streamer> with data-field="marketCap"
  const regex = /<fin-streamer[^>]*data-field="marketCap"[^>]*>\s*([\d\.]+[MBT]?B?)\s*<\/fin-streamer>/;
  const marketCapMatch = contentText.match(regex);
  if (marketCapMatch && marketCapMatch[1]) {
    return marketCapMatch[1];  // Return the captured market cap
  }
  return 'Market Cap not found';  // Return a default message if no market cap is captured
}
// TEST:
console.log("AAPL Market Cap: " + getMarketCap("AAPL"));

// RETURNS BETA (5y Monthly):
function getBeta(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const regex = /<span class="label svelte-tx3nkj">Beta \(5Y Monthly\)<\/span>\s*<span class="value svelte-tx3nkj">([^<]+)<\/span>/;
  const marketCapMatch = contentText.match(regex);
  if (marketCapMatch && marketCapMatch[1]) {
    return marketCapMatch[1];  
  }
  return 'NA';  
}
// TEST:
console.log("GOOG Beta: " + getBeta("GOOG"));


// PE Ratio (TTM):
function getPER(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const regex = /<span class="label svelte-tx3nkj">PE Ratio \(TTM\)<\/span>\s*<span class="value svelte-tx3nkj">\s*<fin-streamer[^>]*>\s*([\d\.]+)\s*<\/fin-streamer>\s*<\/span>/;
  const marketCapMatch = contentText.match(regex);
  if (marketCapMatch && marketCapMatch[1]) {
    return marketCapMatch[1];  
  }
  return 'NA';  
}
// TEST:
console.log("MSFT PER: " + getPER("MSFT"));



// Pending fix after latest change in Yahoo Finance:
function YFgrowth(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}/analysis`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();
  const growth5yr = contentText.match(/Past 5 Years \(per annum\)<\/span><\/td><td class="Ta\(end\) Py\(10px\)">(-?\d+\.\d+%)/); // Pending to fix
  return growth5yr ? growth5yr[1] : "NA";
}
console.log("IBM 5 yr growth: " + YFgrowth("IBM"))

function getThirdFinColValue(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}/balance-sheet`;
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const contentText = res.getContentText();  
  // Match all instances of the pattern, then select the third one
  const pattern = /<div class="[^"]*" data-test="fin-col"><span>([\d,]+)<\/span><\/div>/g; // Pending to fix
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
