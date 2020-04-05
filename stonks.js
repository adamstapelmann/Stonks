/**
 * Runs when the spreadsheet is opened
 */
function onOpen() {
  // Set up Bought headers
  var boughtSheet = SpreadsheetApp.getActive().getSheetByName("Bought");
  var headers = [
    'Ticker',
    'Num Shares',
    'Purchase Price',
    'Date'];
  boughtSheet.getRange('A1:D1').setValues([headers]).setFontWeight('bold');
  boughtSheet.setFrozenRows(1);
  
  // Set up Sold headers
  var soldSheet = SpreadsheetApp.getActive().getSheetByName("Sold");
  var headers = [
    'Ticker',
    'Num Shares',
    'Sale Price',
    'Date'];
  soldSheet.getRange('A1:D1').setValues([headers]).setFontWeight('bold');
  soldSheet.setFrozenRows(1);
  
  // Set up dashboard headers
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var headers = [
    'Ticker',
    'Num Shares',
    'Cost Basis',
    'Quote',
    'Market Value',
    '$ Gain',
    '% Gain'];
  dashboardSheet.getRange('A1:G1').setValues([headers]).setFontWeight('bold');
  dashboardSheet.setFrozenRows(1);

  // Set dashboard data
  setDashboardTickers();
  setDashboardNumShares();
  setDashboardCostBases();
  setDashboardQuotes();
  setDashboardMarketValues();
  setDashboardDollarChange();
  setDashboardPercentChange();
}

// Generates a list of all used tickers by Bought and Sold sheets
function getTickerList() {
  // Get input and output arrays
  var boughtSheet = SpreadsheetApp.getActive().getSheetByName("Bought");
  var boughtTickers = boughtSheet.getDataRange().getValues();  
  var tickers = [];
  
  // Add non-duplicates from boughtTickers to tickers
  // (Sold tickers should be a subset of bought tickers, so not necessary to include them here.)
  for (var i = 1; i < boughtTickers.length; i++) {
    var ticker = boughtTickers[i][0];
    var duplicate = false;
    for (var j = 0; j < tickers.length; j++) {
      if (tickers[j] == ticker) { duplicate = true; }
    }
    if (!duplicate) { tickers.push(ticker); }
  }

  return tickers;
}

// Generates a list of the number of shares for each ticker
function getNumShares() {
  // Access our bought and sold sheets
  var boughtSheet = SpreadsheetApp.getActive().getSheetByName("Bought");
  var boughtValues = boughtSheet.getDataRange().getValues();  
  var soldSheet = SpreadsheetApp.getActive().getSheetByName("Sold");
  var soldValues = soldSheet.getDataRange().getValues();  
  
  var numShares = [];
  var tickers = getTickerList();
  
  // For every ticker, check num shares in our bought and sold lists
  for (var i = 0; i < tickers.length; i++) {
    var ticker = tickers[i];
    numShares[i] = 0;
    
    // Add number of bought shares
    for (var j = 1; j < boughtValues.length; j++) {
      if (boughtValues[j][0] == ticker) {
        numShares[i] = numShares[i] + boughtValues[j][1];
      }
    }
    
    // Subtract number of sold shares
    for (var j = 1; j < soldValues.length; j++) {
      if (soldValues[j][0] == ticker) {
        numShares[i] = numShares[i] - soldValues[j][1];
      }
    }
  }
  
  return numShares;
}

// Generates a list of the cost bases fo each of the tickers of bought stocks
function getCostBases() {
  // Access our bought and sold sheets
  var boughtSheet = SpreadsheetApp.getActive().getSheetByName("Bought");
  var boughtValues = boughtSheet.getDataRange().getValues();  
  var soldSheet = SpreadsheetApp.getActive().getSheetByName("Sold");
  var soldValues = soldSheet.getDataRange().getValues(); 
  
  var costBases = [];
  var tickers = getTickerList();
  var quantities = getNumShares();
  
  // Cost Basis = (Total Spent - Total Earned)/Num Shares
  for (var i = 0; i < tickers.length; i++) {
    var ticker = tickers[i];
    var amountSpent = 0;
    var numShares = quantities[i];
    costBases[i] = 0;
    
    // Add number of bought shares
    for (var j = 1; j < boughtValues.length; j++) {
      if (boughtValues[j][0] == ticker) {
        amountSpent = amountSpent + boughtValues[j][1]*boughtValues[j][2];
      }
    }
    
    // Subtract number of sold shares
    for (var j = 1; j < soldValues.length; j++) {
      if (soldValues[j][0] == ticker) {
        amountSpent = amountSpent - soldValues[j][1]*soldValues[j][2];
      }
    }
    
    costBases[i] = amountSpent/numShares;
  }
  
  return costBases;
}

// Generates a list of the market values for the sum of all shares
function getMarketValues() {
  // Access our bought and sold sheets
  var boughtSheet = SpreadsheetApp.getActive().getSheetByName("Bought");
  var boughtValues = boughtSheet.getDataRange().getValues();  
  var soldSheet = SpreadsheetApp.getActive().getSheetByName("Sold");
  var soldValues = soldSheet.getDataRange().getValues();
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var dashboardValues = dashboardSheet.getDataRange().getValues(); 
  
  var tickers = getTickerList();
  var numShares = col2row(dashboardSheet.getRange(2, 2, tickers.length, 1).getValues())[0];
  var quotes = col2row(dashboardSheet.getRange(2, 4, tickers.length, 1).getValues())[0];
  var marketValues = [];
  
  // market value = num shares * quote
  for (var i = 0; i < tickers.length; i++) {
    marketValues[i] = numShares[i]*quotes[i]; 
  }
  
  return marketValues;
}

// Generates a list of the dollar gain/loss for each ticker
function getDollarChange() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var dashboardValues = dashboardSheet.getDataRange().getValues(); 
  
  var tickers = getTickerList();
  var numShares = col2row(dashboardSheet.getRange(2, 2, tickers.length, 1).getValues())[0];
  var costBases = col2row(dashboardSheet.getRange(2, 3, tickers.length, 1).getValues())[0];
  var quotes = col2row(dashboardSheet.getRange(2, 4, tickers.length, 1).getValues())[0];
  var dollarChange = [];
  
  // $ change = num shares * (quote - cost basis)
  for (var i = 0; i < tickers.length; i++) {
    dollarChange[i] = numShares[i]*(quotes[i]-costBases[i]); 
  }
  
  return dollarChange;
}

// Generates a list of the percent gain/loss for each ticker
function getPercentChange() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var dashboardValues = dashboardSheet.getDataRange().getValues(); 
  
  var tickers = getTickerList();
  var costBases = col2row(dashboardSheet.getRange(2, 3, tickers.length, 1).getValues())[0];
  var quotes = col2row(dashboardSheet.getRange(2, 4, tickers.length, 1).getValues())[0];
  var percentChange = [];
  
  // % change = 100 * (quote/cost basis) - 100
  for (var i = 0; i < tickers.length; i++) {
    percentChange[i] = 100*(quotes[i]/costBases[i]) - 100; 
  }
  
  return percentChange;
}

// Add tickers to dashboard (A2:A?)
function setDashboardTickers() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var tickers = getTickerList();
  dashboardSheet.getRange(2, 1, tickers.length, 1).setValues(row2col([tickers]));
}

// Add num shares to dashboard (B2:B?)
function setDashboardNumShares() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var numShares = getNumShares();
  dashboardSheet.getRange(2, 2, numShares.length, 1).setValues(row2col([numShares]));
}

// Add cost bases to dashboard (C2:C?)
function setDashboardCostBases() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var costBases = getCostBases();
  dashboardSheet.getRange(2, 3, costBases.length, 1).setValues(row2col([costBases]));
  dashboardSheet.getRange(2, 3, costBases.length, 1).setNumberFormat('0.00');
}

// Add quotes to dashboard (D2:D?)
function setDashboardQuotes() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var tickers = getTickerList();
  dashboardSheet.getRange(2, 4, tickers.length, 1).setFormulaR1C1('=GOOGLEFINANCE(R[0]C[-3])')
}

// Add market values to dashboard (E2:E?)
function setDashboardMarketValues() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var marketValues = getMarketValues();
  dashboardSheet.getRange(2, 5, marketValues.length, 1).setValues(row2col([marketValues]));
}

// Add market values to dashboard (F2:E?)
function setDashboardDollarChange() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var dollarChange = getDollarChange();
  dashboardSheet.getRange(2, 6, dollarChange.length, 1).setValues(row2col([dollarChange]));
  dashboardSheet.getRange(2, 6, dollarChange.length, 1).setNumberFormat('0.00');
  
  
  // Add colors
  var row = col2row(dashboardSheet.getRange(2, 6, dollarChange.length, 1).getValues())[0];
  for (var i = 0; i < row.length; i++) {
    if (row[i] < 0) {
      dashboardSheet.getRange(2 + i, 6, 1, 1).setBackgroundRGB(255, 120, 120);
    }
    if (row[i] >= 0) {
      dashboardSheet.getRange(2 + i, 6, 1, 1).setBackgroundRGB(120, 255, 120);
    }
  }
}

// Add market values to dashboard (G2:E?)
function setDashboardPercentChange() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var percentChange = getPercentChange();
  dashboardSheet.getRange(2, 7, percentChange.length, 1).setValues(row2col([percentChange]));
  dashboardSheet.getRange(2, 7, percentChange.length, 1).setNumberFormat('0.00');
  
  // Add colors
  var row = col2row(dashboardSheet.getRange(2, 7, percentChange.length, 1).getValues())[0];
  for (var i = 0; i < row.length; i++) {
    if (row[i] < 0) {
      dashboardSheet.getRange(2 + i, 7, 1, 1).setBackgroundRGB(255, 120, 120);
    }
    if (row[i] >= 0) {
      dashboardSheet.getRange(2 + i, 7, 1, 1).setBackgroundRGB(120, 255, 120);
    }
  }  
}

// Sets a row of totals on the dashboard
function setDashboardTotals() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
  var tickers = getTickerList();
  dashboardSheet.getRange(tickers.length + 2, 1, 1, 1).setValues([["Total"]]).setFontWeight('bold');
  
  // Get totals
  var totalGain = 0;
  var totalSpent = 0;
  var numShares = col2row(dashboardSheet.getRange(2, 2, tickers.length, 1).getValues())[0];
  var costBases = col2row(dashboardSheet.getRange(2, 3, tickers.length, 1).getValues())[0];
  var gainLoss = col2row(dashboardSheet.getRange(2, 6, tickers.length, 1).getValues())[0];
  for (var i = 0; i < gainLoss.length; i++) {
    totalGain = totalGain + gainLoss[i];
    totalSpent = totalSpent + (numShares[i]*costBases[i]);
  }
    
  // Set dollar change total
  dashboardSheet.getRange(tickers.length + 2, 6, 1, 1).setValues([[totalGain]]).setFontWeight('bold').setNumberFormat('0.00');
  if (totalGain < 0) {
    dashboardSheet.getRange(tickers.length + 2, 6, 1, 1).setBackgroundRGB(255, 120, 120);
  }
  
  if (totalGain >= 0) {
    dashboardSheet.getRange(tickers.length + 2, 6, 1, 1).setBackgroundRGB(120, 255, 120);
  }
  
  // Set percent change total
  var percentGain = 100*(totalGain / totalSpent);
  dashboardSheet.getRange(tickers.length + 2, 7, 1, 1).setValues([[percentGain]]).setFontWeight('bold').setNumberFormat('0.00');
  if (totalGain < 0) {
    dashboardSheet.getRange(tickers.length + 2, 7, 1, 1).setBackgroundRGB(255, 120, 120);
  }
  
  if (totalGain >= 0) {
    dashboardSheet.getRange(tickers.length + 2, 7, 1, 1).setBackgroundRGB(120, 255, 120);
  }
  
}

// API call to get current bitcoin price
function getBtcPrice() {
  var response = UrlFetchApp.fetch("https://api.coindesk.com/v1/bpi/currentprice.json").getContentText();
  var jsonObj = JSON.parse(response);
  var rate = jsonObj.bpi.USD.rate;
  return rate;
}

// Maps a column to a row
function col2row(column) {
  return [column.map(function(row) {return row[0];})];
} 

// Maps a row to a column
function row2col(row) {
  return row[0].map(function(elem) {return [elem];});
}