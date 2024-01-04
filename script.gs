const API_BASE_URL = "https://api.coinpaprika.com/v1";
const API_PRO_BASE_URL = "https://api-pro.coinpaprika.com/v1";
const API_KEY_CELL = "A1";

function onOpen() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("config");
  const apiKeyCell = configSheet ? configSheet.getRange(API_KEY_CELL).getValue() : null;
  SpreadsheetApp.getUi()
      .createMenu(`CoinPaprika ${apiKeyCell ? 'Pro' : ''}`)
      .addItem('Update All Data', 'CRYPTODATAJSON')
      .addItem('Get Specific Data', 'CRYPTODATA_UI')
      .addItem('Get Historical Data', 'CRYPTODATAHISTORY_UI')
      .addItem('Get Coin Details', 'CRYPTODATACOINDETAILS_UI')
      .addToUi();
}

function CRYPTODATA_UI() {
  var ui = SpreadsheetApp.getUi();

  var symbolResponse = ui.prompt('Enter the Symbol:');
  if (symbolResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a symbol.');
    return;
  }
  var symbol = symbolResponse.getResponseText();

  var colNameResponse = ui.prompt('Enter the Column Name:');
  if (colNameResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a column name.');
    return;
  }
  var colName = colNameResponse.getResponseText();

  var result = CRYPTODATA(symbol, colName);

  // Display the result
  ui.alert('Result: ' + result);
}


function CRYPTODATA(symbol, colName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  const data = sheet.getDataRange().getValues();
  const colIndex = data[0].indexOf(colName);

  if (colIndex === -1) {
    return "datatype not found";
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === symbol) {
      return data[i][colIndex];
    }
  }

  return "symbol not found";
}

function CRYPTODATAHISTORY_UI() {
  var ui = SpreadsheetApp.getUi();

  var symbolResponse = ui.prompt('Enter the Symbol:');
  if (symbolResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a symbol.');
    return;
  }
  var symbol = symbolResponse.getResponseText();

  var dateResponse = ui.prompt('Enter the Date (YYYY-MM-DD):');
  if (dateResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a date.');
    return;
  }
  var date = dateResponse.getResponseText();

  var typeResponse = ui.prompt('Enter the Type (e.g., "price"):');
  if (typeResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a type.');
    return;
  }
  var type = typeResponse.getResponseText();

  var quoteResponse = ui.prompt('Enter the Quote Currency (default: "usd"):');
  var quote = quoteResponse.getSelectedButton() == ui.Button.OK ? quoteResponse.getResponseText() : "usd";

  var result = CRYPTODATAHISTORY(symbol, date, type, quote);

  ui.alert('Result: ' + result);
}


function CRYPTODATAHISTORY(symbol, date, type, quote) {
  quote = quote || "usd";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  const data = sheet.getDataRange().getValues();
  const colIndex = data[0].indexOf("id");

  if (colIndex === -1) {
    return "datatype not found";
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === symbol) {
      const coinId = data[i][colIndex];
      const formattedDate = date instanceof Date ? date.toISOString() : date;
      const url = fetchUrl(`tickers/${coinId}/historical?start=${formattedDate}&interval=5m&limit=1&quote=${quote}`);

      const responseData = fetchUrl(url);
      return responseData[0].hasOwnProperty(type)
          ? responseData[0][type]
          : "price not found";
    }
  }

  return "symbol not found";
}

function CRYPTODATACOINDETAILS_UI() {
  var ui = SpreadsheetApp.getUi();

  var coinIdResponse = ui.prompt('Enter the Coin ID:');
  if (coinIdResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a coin ID.');
    return;
  }
  var coinId = coinIdResponse.getResponseText();

  var colNameResponse = ui.prompt('Enter the Column Name:');
  if (colNameResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('You did not enter a column name.');
    return;
  }
  var colName = colNameResponse.getResponseText();

  var result = CRYPTODATACOINDETAILS(coinId, colName);

  // Display the result
  ui.alert('Result: ' + result);
}


function CRYPTODATACOINDETAILS(coin_id, colName) {
  const param = encodeURI(coin_id);

  if (colName.indexOf("/") !== -1) {
    const [quote, property] = colName.split("/");
    const url = fetchUrl(`tickers/${param}?quotes=${quote}`);

    const responseData = fetchUrl(url);

    if (
        responseData.hasOwnProperty("error") &&
        responseData["error"] === "id not found"
    ) {
      return "coin_id not found";
    }

    if (responseData["quotes"][quote].hasOwnProperty(property)) {
      return responseData["quotes"][quote][property];
    } else {
      return "property not found";
    }
  } else {
    const url = fetchUrl(`coins/${param}`);
    const responseData = fetchUrl(url);

    if (
        responseData.hasOwnProperty("error") &&
        responseData["error"] === "id not found"
    ) {
      return "coin_id not found";
    }

    if (responseData.hasOwnProperty(colName)) {
      return responseData[colName];
    } else {
      return "property not found";
    }
  }
}

function CRYPTODATAJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "data";
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  const ui = SpreadsheetApp.getUi();

  const url = "tickers?quotes=USD,BTC";
  const dataAll = fetchUrl(url);
  const dataSet = dataAll.sort((a, b) => (a.rank > b.rank ? 1 : -1));

  const rows = [headers];

  dataSet.forEach((data) => {
    const rowData = headers.map((header) => {
      if (header.startsWith("btc_")) {
        const key = header.slice(4);
        return data.quotes["BTC"][key] || "";
      }
      if (header.startsWith("usd_")) {
        const key = header.slice(4);
        return data.quotes["USD"][key] || "";
      }
      return data[header] || "";
    });
    rows.push(rowData);
  });

  const dataRange = sheet.getRange(1, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}

/* UTILS:
 ***********/

function fetchUrl(endpoint) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("config");
  const apiKey = configSheet ? configSheet.getRange(API_KEY_CELL).getValue() : null;
  const reqHeaders = apiKey ? { Authorization: `${apiKey}` } : {};
  const baseUrl = apiKey ? API_PRO_BASE_URL : API_BASE_URL;
  const url = `${baseUrl}/${endpoint}`;

  const response = UrlFetchApp.fetch(url, { reqHeaders, muteHttpExceptions: true });
  RESPONSECODE(response);
  return JSON.parse(response.getContentText());
}

function RESPONSECODE(v) {
  const responseCode = v.getResponseCode();

  if (responseCode === 429) {
    throw new Error("Too many requests");
  } else if (responseCode !== 200) {
    throw new Error(`Server error. Response: ${v.getContentText()}`);
  }
}

/* CRYPTODATAJSON() - DATA HEADERS
 **********************************/

const headers = [
  "id", "name", "symbol", "rank", "circulating_supply", "total_supply", "max_supply",
  "beta_value", "first_data_at", "last_updated", "btc_price", "btc_volume_24h",
  "btc_volume_24h_change_24h", "btc_market_cap", "btc_market_cap_change_24h",
  "btc_percent_change_15m", "btc_percent_change_30m", "btc_percent_change_1h",
  "btc_percent_change_6h", "btc_percent_change_12h", "btc_percent_change_24h",
  "btc_percent_change_7d", "btc_percent_change_30d", "btc_percent_change_1y",
  "btc_ath_price", "btc_ath_date", "btc_percent_from_price_ath", "usd_price",
  "usd_volume_24h", "usd_volume_24h_change_24h", "usd_market_cap",
  "usd_market_cap_change_24h", "usd_percent_change_15m", "usd_percent_change_30m",
  "usd_percent_change_1h", "usd_percent_change_6h", "usd_percent_change_12h",
  "usd_percent_change_24h", "usd_percent_change_7d", "usd_percent_change_30d",
  "usd_percent_change_1y", "usd_ath_price", "usd_ath_date",
  "usd_percent_from_price_ath",
];
