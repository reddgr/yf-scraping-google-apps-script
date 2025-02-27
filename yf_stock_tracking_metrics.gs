// RETURNS LAST CLOSE PRICE:
function getPreviousClose(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const contentText = res.getContentText();
  
  // Regex to extract previous close value from the <fin-streamer> with data-field="regularMarketPreviousClose"
  const regex = /<fin-streamer[^>]*data-field="regularMarketPreviousClose"[^>]*>\s*([\d\.]+)\s*<\/fin-streamer>/;
  const match = contentText.match(regex);
  
  if (match && match[1]) {
    return match[1]; // Return the captured previous close value
  }
  return 'Previous Close not found'; // Return a default message if not found
}

// TEST:
const test_ticker = "AAPL"
console.log(`${test_ticker} previous close: ${getPreviousClose(test_ticker)}`);

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
console.log(`${test_ticker} market cap: ${getMarketCap(test_ticker)}`);

// RETURNS BETA (5y Monthly):
function getBeta(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const contentText = res.getContentText();
  
  // Regex to locate "Beta (5Y Monthly)" and extract the following number
  const regex = /<span[^>]*title="Beta \(5Y Monthly\)"[^>]*>.*?<\/span>\s*<span[^>]*class="value[^"]*"[^>]*>\s*([\d\.]+)\s*<\/span>/;
  const match = contentText.match(regex);
  
  if (match && match[1]) {
    return match[1]; // Return the captured Beta value
  }
  return 'Beta (5Y Monthly) not found'; // Default message if not found
}
// TEST:
console.log(`AAPL Beta (5Y Monthly): ${getBeta("AAPL")}`);

function getPER(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const contentText = res.getContentText();
  
  // Updated regex to extract the PE Ratio (TTM) value from the fin-streamer tag
  const regex = /<span[^>]*title="PE Ratio \(TTM\)"[^>]*>.*?<\/span>\s*<span[^>]*class="value[^"]*"[^>]*>\s*<fin-streamer[^>]*data-value="([\d\.]+)"/;
  const match = contentText.match(regex);
  
  if (match && match[1]) {
    return match[1]; // Return the extracted PE Ratio value
  }
  return 'PE Ratio (TTM) not found'; // Default message if not found
}

// TEST:
console.log(`AAPL PE Ratio (TTM): ${getPER("AAPL")}`);

function getPEG(ticker) {
  const url = `https://finance.yahoo.com/quote/${ticker}/key-statistics`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const contentText = res.getContentText();
  // Regex to match the first number in the PEG Ratio (5yr expected) row
  const regex = /<tr class="yf-kbx2lo alt">\s*<td class="yf-kbx2lo">PEG Ratio \(5yr expected\)<\/td>\s*<td class="yf-kbx2lo">([\d\.]+)\s*<\/td>/;
  const pegMatch = contentText.match(regex);
  if (pegMatch && pegMatch[1]) {
    return pegMatch[1];  // Return the captured PEG Ratio value
  }
  return 'PEG Ratio not found';  // Return a default message if no PEG Ratio is captured
}
// TEST:
console.log(`${test_ticker} PEG: ${getPEG(test_ticker)}`);

// DUPLICATE A SHEET WITH VALUES TO KEEP A MANUAL BACKUP OF YAHOO FINANCE METRICS
function duplicateSheet2() {
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







