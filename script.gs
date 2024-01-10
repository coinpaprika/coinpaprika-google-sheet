const API_BASE_URL = "https://api.coinpaprika.com/v1";
const API_PRO_BASE_URL = "https://api-pro.coinpaprika.com/v1";
const API_KEY_CELL = "B1";

/**
 * Retrieves specific data for a given cryptocurrency symbol from a Google Sheet named 'data'.
 *
 * This function searches for the given symbol in the 'data' sheet and returns the value from the specified column name.
 * If the column name is not found, it returns 'datatype not found'. If the symbol is not found, it returns 'symbol not found'.
 * The data sheet is expected to have the symbol in the third column (index 2).
 *
 * @param {string} symbol - The symbol of the cryptocurrency (e.g., 'BTC').
 * @param {string} colName - The name of the column from which to retrieve data (e.g., 'price').
 * @returns {string|number} - Returns the data from the specified column for the given symbol, or an error message if not found.
 * @customfunction
 */
function CP_TICKERS(symbol, colName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

/**
 * Fetches historical data for a given cryptocurrency symbol.
 *
 * This function looks up the coin ID based on the provided symbol from a sheet named 'data',
 * then makes an API call to fetch historical data for that coin. The historical data is
 * retrieved for the specified date and quote currency. If the requested type of data (e.g., 'price')
 * is available in the response, it is returned; otherwise, a 'price not found' message is returned.
 *
 * @param {string} coinId - The ID of the cryptocurrency (e.g., 'BTC').
 * @param {string} date - The date for which historical data is required, in 'YYYY-MM-DD' format.
 * @param {string} type - The type of historical data to retrieve (e.g., 'price').
 * @param {string} [quote='usd'] - The quote currency (default is 'usd').
 * @param {string} apiKey - The pro API key, empty string for free plan.
 * @returns {string|number} - Returns the requested historical data or an error message if not found.
 * @customfunction
 */
function CP_TICKERS_HISTORY(coinId, date, type, quote, apiKey) {
  quote = quote || "usd";

  const formattedDate = date instanceof Date ? date.toISOString() : date;
  const url = FETCHURL_(`tickers/${coinId}/historical?start=${formattedDate}&interval=5m&limit=1&quote=${quote}`);

  const responseData = FETCHURL_(url, apiKey);
  return responseData[0].hasOwnProperty(type)
      ? responseData[0][type]
      : "price not found";
}

/**
 * Retrieves specific details about a cryptocurrency based on its coin ID.
 *
 * This function uses the coin_id to fetch data about the cryptocurrency. If colName contains a slash ('/'),
 * it is split to derive quote and property, then a specific API endpoint is used to fetch the data.
 * If colName does not contain a slash, a different endpoint is used. The function returns either the
 * specified property of the cryptocurrency or an error message if the property or coin ID is not found.
 * @param {string} coin_id - The unique identifier of the cryptocurrency (e.g., 'bitcoin').
 * @param {string} colName - The name of the column (property) to retrieve, which can be a simple name or
 *                           in the format 'quote/property' (e.g., 'usd/price').
 * @param {string} apiKey - The pro API key, empty string for free plan.
 * @returns {string|number} - Returns the requested property value of the cryptocurrency, or an error message
 *                            if the property or coin ID is not found.
 *
 * @customfunction
 */
function CP_COINS(coin_id, colName, apiKey) {
  const param = encodeURI(coin_id);

  if (colName.indexOf("/") !== -1) {
    const [quote, property] = colName.split("/");
    const url = FETCHURL_(`tickers/${param}?quotes=${quote}`);

    const responseData = FETCHURL_(url, apiKey);

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
    const url = FETCHURL_(`coins/${param}`);
    const responseData = FETCHURL_(url, apiKey);

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

/**
 * Fetches and populates cells with global cryptocurrency data.
 *
 * This function retrieves global cryptocurrency data (like market cap, volume, etc.) and returns it in a two-dimensional array format.
 * The first row of the array contains the headers (data keys), and the second row contains the corresponding values.
 * This function is designed to be used as a custom function in Google Sheets and will populate the cells with data starting from the cell it's entered in.
 * @param {string} apiKey - The pro API key, empty string for free plan.
 * @returns {Array<Array<string|number>>} A two-dimensional array where the first row is headers and the second row is values.
 * @customfunction
 */
function CP_GLOBAL(apiKey) {
  const data = FETCHURL_('global', apiKey);

  const headers = [];
  const values = [];

  for (let key in data) {
    if (data.hasOwnProperty(key)) {
      headers.push(key);
      values.push(data[key]);
    }
  }

  return [headers, values];
}

/**
 * Fetches cryptocurrency data and returns an array of arrays containing headers and values.
 *
 * This function retrieves a list of cryptocurrency tickers with their respective data in USD and BTC quotes,
 * sorts them by rank, and returns an array of arrays. The first array contains the column names (headers),
 * and the subsequent arrays contain the corresponding values for each cryptocurrency.
 *
 * @returns {Array<Array<string|number>>} An array of arrays where the first array is headers and the subsequent arrays are values.
 * @customfunction
 */
function CP_TICKERS_ALL(apiKey) {
  const url = "tickers?quotes=USD,BTC";
  const dataAll = FETCHURL_(url, apiKey);
  const dataSet = dataAll.sort((a, b) => (a.rank > b.rank ? 1 : -1));

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

  const values = dataSet.map(data => headers.map(header => {
    if (header.startsWith("btc_")) {
      const key = header.slice(4);
      return data.quotes["BTC"][key] || "";
    }
    if (header.startsWith("usd_")) {
      const key = header.slice(4);
      return data.quotes["USD"][key] || "";
    }
    return data[header] || "";
  }));

  values.unshift(headers);

  return values;
}


/* UTILS:
 ***********/

function FETCHURL_(endpoint, apiKey) {
  if (apiKey && apiKey.length > 0 && !apiKey.startsWith("api_")) {
    throw new Error("Invalid API key. The key should start with 'api_'.");
  }

  const reqHeaders = apiKey ? { Authorization: `${apiKey}` } : {};
  const baseUrl = apiKey ? API_PRO_BASE_URL : API_BASE_URL;
  const url = `${baseUrl}/${endpoint}`;

  const response = UrlFetchApp.fetch(url, { headers:reqHeaders, muteHttpExceptions: true });
  RESPONSECODE_(response);
  return JSON.parse(response.getContentText());
}

function RESPONSECODE_(v) {
  const responseCode = v.getResponseCode();

  if (responseCode === 429) {
    throw new Error("Too many requests");
  } else if (responseCode !== 200) {
    throw new Error(`Server error. Response: ${v.getContentText()}`);
  }
}
