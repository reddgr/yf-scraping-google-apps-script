# Google Sheets UrlFetchApp Example Functions

This repository contains a set of example functions for educational purposes, demonstrating how to use the Google Apps Script `UrlFetchApp` class to integrate financial metrics from Yahoo Finance into Google Sheets. These functions are intended to serve as illustrative examples, providing a basic framework for fetching real-time financial data for personal learning and exploration.

## Overview of Functions

- **`Yfinance(ticker)`**: Demonstrates fetching the current stock price for a specified ticker.
- **`YFmc(ticker)`**: Shows how to retrieve the market capitalization of a stock.
- **`YFbeta(ticker)`**: Example of obtaining the beta coefficient to indicate stock volatility.
- **`YFper(ticker)`**: Illustrates returning the Price to Earnings (P/E) ratio.
- **`YFgrowth(ticker)`**: Provides an example of estimating the five-year growth rate.
- **`getThirdFinColValue(ticker)`**: Extracts a specific value from a stock's balance sheet as an example.
- **`extractGrossMargin(ticker)`**: Scrapes the gross margin percentage. (URL adjustment to Yahoo Finance is required.)
- **`duplicateSheet()`**: Example function to duplicate a specified sheet within the same spreadsheet, useful for data backup.

These functions utilize the Google Apps Script `UrlFetchApp` class to make HTTP requests to Yahoo Finance's web pages and parse the returned HTML for specific financial data. They are designed to be copied into a Google Apps Script project linked to a Google Sheet for direct application and experimentation.

## Implementation for Educational Use

1. Open the Google Sheet you wish to experiment with.
2. Go to `Extensions > Apps Script`, and start a new script project.
3. Copy and paste the example functions into the script editor.
4. Save the script project with a meaningful name.
5. In your Google Sheet, call the functions within cells using the syntax `=FunctionName("Ticker")`, replacing `FunctionName` with any of the provided examples and `"Ticker"` with a valid stock ticker symbol.

## Important Notes

- These example functions employ the `UrlFetchApp` class, documented [here](https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app), for educational purposes.
- As illustrative examples, they may require modifications to work with current Yahoo Finance web page structures or to meet specific educational objectives.
- This set of functions is intended for personal learning and should not be used for commercial purposes without understanding the implications, including adherence to Yahoo Finance's terms of service.

## Disclaimer

These functions are provided as-is for educational purposes to demonstrate coding practices and techniques with Google Apps Script. They are not intended for real-world financial analysis or investment decisions. Users should exercise caution and verify all data independently.

---

This README clarifies the intent and usage of the Google Sheets Yahoo Finance Example Functions, highlighting their educational purpose and guiding users on how to experiment with them for personal learning.

This README was generated with JavaScript Code Streamliner (https://chat.openai.com/g/g-htqjs4zMj-javascript-code-streamliner).

