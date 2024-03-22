```
CoinPaprika Google Sheets Integration
```

Welcome to the updated CoinPaprika Google Sheets Integration! This tool integrates real-time cryptocurrency data from CoinPaprika into Google Sheets, enabling advanced tracking and analysis of the cryptocurrency market.

## Features

- **Direct API Integration:** Fetch real-time cryptocurrency data from CoinPaprika's API.
- **Custom Functions:** Access various data points, including historical data, ticker information, global crypto market data, and specific coin details.
- **API-PRO Compatibility:** Leverage enhanced features with CoinPaprika's API-PRO for more detailed data access.
- **Simple Configuration:** Use API key as a parameter for expanded data access.

## Setup Guide

### Step 1: Open Google Sheets

[Create](https://sheet.new/) or [open](https://drive.google.com/) an existing Google Sheets document for cryptocurrency data integration.
![Step 1](screenshots/step1.png)

### Step 2: Access Script Editor

In Google Sheets, navigate to "Extensions" > "Apps Script" to open the Script Editor.
![Step 2](screenshots/step2.png)

### Step 3: Copy and Paste the Script
Copy the CoinPaprika integration [script from the repository](https://raw.githubusercontent.com/coinpaprika/coinpaprika-google-sheet/main/script.gs) and paste it into the Script Editor.
![Step 3](screenshots/step3.png)

### Step 4: Save and Authorize

Save the script. If running for the first time, Google will prompt you to authorize the script.
![Step 4-1](screenshots/step4-1.png)
![Step 4-2](screenshots/step4-2.png)

### Step 5: Start Using Functions

Use the custom functions in your sheet with the required parameters, including the API key if you have one.

## Available Functions

1. **CP_TICKERS(coinId, apiKey)**:
   - Fetches specific data for a given cryptocurrency symbol.
   - Parameters:
      - `coinId` (string): Symbol of the cryptocurrency (e.g., 'BTC').
      - `apiKey` (string): API key for CoinPaprika, if available.
   - Formula Usage example: 
      - API pro - without limits:
     ```excel
     =CP_TICKERS("btc-bitcoin", "your_api_key_here")
     ```
      - API free:
     ```excel
     =CP_TICKERS("btc-bitcoin")
     ```

2. **CP_TICKERS_HISTORY(coinId, date, interval, limit, quote, apiKey)**:
   - Retrieves historical data for a specified cryptocurrency.
   - Parameters:
      - `coinId` (string): ID of the cryptocurrency (e.g., 'btc-bitcoin').
      - `date` (string): Date for historical data in 'YYYY-MM-DD' format.
      - `interval` (string): Interval for historical data (e.g., '1h').
      - `limit` (string): Limit of historical data to retrieve.
      - `quote` (string): Quote currency (default is 'usd').
      - `apiKey` (string): API key for CoinPaprika, if available.
   - Formula Usage Example:
     ```excel
     =CP_TICKERS_HISTORY("btc-bitcoin", "2022-01-01", "1h", "10", "usd", "your_api_key_here")
     ```

3. **CP_COINS(coin_id, apiKey)**:
   - Retrieves specific details about a cryptocurrency based on its coin ID.
   - Parameters:
      - `coin_id` (string): Unique identifier of the cryptocurrency (e.g., 'bitcoin').
      - `apiKey` (string): API key for CoinPaprika, if available.
   - Formula Usage Example:
      - API pro - without limits:
     ```excel
     =CP_COINS("btc-bitcoin", "your_api_key_here")
     ```
     - API free:
     ```excel
     =CP_COINS("btc-bitcoin")
     ```

4. **CP_GLOBAL(apiKey)**:
   - Fetches and populates cells with global cryptocurrency data.
   - Parameters:
      - `apiKey` (string): API key for CoinPaprika, if available.
   - Formula Usage Example:
      - API pro - without limits:
     ```excel
     =CP_GLOBAL("your_api_key_here")
     ```
     - API free:
     ```excel
     =CP_GLOBAL()
     ```

5. **CP(coinId, apiKey)**:
    - Fetches USD price for a given cryptocurrency symbol.
    - Parameters:
        - `coinId` (string): Symbol of the cryptocurrency (e.g., 'btc-bitcoin').
        - `apiKey` (string): API key for CoinPaprika, if available.
    - Formula Usage example:
        - API pro - without limits:
      ```excel
      =CP("btc-bitcoin", "your_api_key_here")
      ```
        - API free:
      ```excel
      =CP("btc-bitcoin")
      ```

## API-PRO Features

Upgrade to CoinPaprika API-PRO for additional insights:
- **Enhanced Historical Data**: Access detailed historical data.
- **Wider Asset Coverage**: Track a broader range of cryptocurrencies.
- **Frequent Price Updates**: Get more refined price updates.

[Upgrade to API-PRO](https://coinpaprika.com/api) for advanced features.

## About CoinPaprika

[CoinPaprika](https://coinpaprika.com) provides comprehensive market data and is a trusted resource for cryptocurrency insights.

## Support and Feedback

For support, suggestions, or feedback, please visit [GitHub Issues](https://github.com/your-repo-link).
