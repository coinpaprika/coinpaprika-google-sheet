const API_BASE_URL = "https://api.coinpaprika.com/v1";
const API_PRO_BASE_URL = "https://api-pro.coinpaprika.com/v1";

/**
 /**
 * Retrieves price(in USD) of a cryptocurrency based on its coin ID.
 *
 * This function uses the coin_id to fetch data about the cryptocurrency.
 *
 * @param {string} coin_id - The unique identifier of the cryptocurrency (e.g., 'bitcoin').
 * @param {string} apiKey - The API key required for access.
 * @returns {number} - Returns price.
 * @customfunction
 */
function CP(coin_id, apiKey) {
  coin_id='btc-bitcoin'
  const url = `tickers/${coin_id}?quotes=USD,BTC`;
  try {
    const data = FETCHURL_(url, apiKey);
    let key = "price"
    return data.quotes["USD"][key] || "";
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Retrieves specific data for a given cryptocurrency symbol.
 *
 * This function makes an API call to fetch data for a specific cryptocurrency symbol, including
 * details like price, market cap, and more.
 *
 * @param {string} coinId - The symbol of the cryptocurrency (e.g., 'BTC').
 * @param {string} apiKey - The API key required for access.
 * @returns {Array<Array<string|number>>} - Returns an array with headers as the first row and values
 *                                           as the second row, or an error message if data is not found.
 * @customfunction
 */
function CP_TICKERS(coinId, apiKey) {
  const url = `tickers/${coinId}?quotes=USD,BTC`;
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
  try {
    const data = FETCHURL_(url, apiKey);
    const values = [headers.map(header => {
      if (header.startsWith("btc_")) {
        const key = header.slice(4);
        return data.quotes["BTC"][key] || "";
      }
      if (header.startsWith("usd_")) {
        const key = header.slice(4);
        return data.quotes["USD"][key] || "";
      }
      return data[header] || "";
    })];

    values.unshift(headers);

    return values;
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Fetches historical data for a given cryptocurrency symbol.
 *
 * This function makes an API call to fetch historical data for a specific cryptocurrency symbol.
 * The historical data includes information such as price, volume, and market cap.
 *
 * @param {string} coinId - The symbol of the cryptocurrency (e.g., 'BTC').
 * @param {string|Date} date - The date for which historical data is required in 'YYYY-MM-DD' format or as a Date object.
 * @param {string} interval - The interval for historical data (e.g., '1h').
 * @param {string} limit - The limit of historical data to retrieve.
 * @param {string} quote - The quote currency (default is 'usd').
 * @param {string} apiKey - The API key required for access.
 * @returns {Array<Array<string|number>>} - Returns an array with headers as the first row and historical data as subsequent rows,
 *                                          or an error message if data is not found.
 * @customfunction
 */
function CP_TICKERS_HISTORY(coinId, date, interval, limit, quote, apiKey) {
  quote = quote || "usd";

  const formattedDate = date instanceof Date ? date.toISOString() : date;
  const url = `tickers/${coinId}/historical?start=${formattedDate}&interval=${interval}&limit=${limit}&quote=${quote}`;
  try {
    const data = FETCHURL_(url, apiKey);
    const result = [];
    const keys = Object.keys(data[0]);
    result.push(keys);

    for (const entry of data) {
      const values = keys.map(key => entry[key] || "");
      result.push(values);
    }

    return result;
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 /**
 * Retrieves specific details about a cryptocurrency based on its coin ID.
 *
 * This function uses the coin_id to fetch data about the cryptocurrency.
 *
 * @param {string} coin_id - The unique identifier of the cryptocurrency (e.g., 'bitcoin').
 * @param {string} apiKey - The API key required for access.
 * @returns {Array<Array<string|number>>} - Returns an array with headers as the first row and cryptocurrency details as subsequent rows,
 *                                          or an error message if data is not found.
 * @customfunction
 */
function CP_COINS(coin_id, apiKey) {
  const param = encodeURI(coin_id);
  const url = `coins/${param}`
  try {
    const data = FETCHURL_(url, apiKey);
    const headers = [];
    const values = [];

    for (let key in data) {
      if (data.hasOwnProperty(key)) {
        headers.push(key);
        values.push(data[key]);
      }
    }

    return [headers, values];
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Fetches and populates cells with global cryptocurrency data.
 *
 * This function retrieves global cryptocurrency data (like market cap, volume, etc.) and returns it in a two-dimensional array format.
 * The first row of the array contains the headers (data keys), and the second row contains the corresponding values.
 * This function is designed to be used as a custom function in Google Sheets and will populate the cells with data starting from the cell it's entered in.
 *
 * @param {string} apiKey - The API key required for access.
 * @returns {Array<Array<string|number>>} - Returns an array with headers as the first row and global cryptocurrency data as subsequent rows,
 *                                          or an error message if data is not found.
 * @customfunction
 */
function CP_GLOBAL(apiKey) {
  const url = 'global'
  try {
    const data = FETCHURL_(url, apiKey);
    const headers = [];
    const values = [];

    for (let key in data) {
      if (data.hasOwnProperty(key)) {
        headers.push(key);
        values.push(data[key]);
      }
    }

    return [headers, values];
  } catch (e) {
    return `Error: ${e.message}`;
  }
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

  if (responseCode === 429 || responseCode === 402) {
    throw new Error("You have reached the free plan limit, please visit https://coinpaprika.com/api for more information");
  } else if (responseCode !== 200) {
    throw new Error(`Server error. Response: ${v.getContentText()}`);
  }
}