# Fear & Greed Index Google Sheets Integration

A Google Apps Script that fetches the latest [CoinMarketCap Fear & Greed Index](https://coinmarketcap.com/charts/fear-and-greed-index/)
into Google Sheets with email alerts for extreme values.

## Features
- Fetch Fear & Greed Index from CoinMarketCap API
- Auto-updates data in Google Sheets
- Email alerts when thresholds are crossed
- Custom menu in Sheets to run/update

## Setup
1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Paste the code from `Code.gs`
4. Replace placeholders with your own API key and email
5. Save and run `fetchLatestFearGreedIndex`
6. Authorize when prompted

## Example
https://docs.google.com/spreadsheets/d/1gRYowKzq6VlB8kFAJ2gVZ02qVfRD3948_OBsJSxLyBw/edit?usp=sharing
![CMCf gdemo](https://github.com/user-attachments/assets/57590dd7-f10a-4c77-9711-3880b3b2588b)
Email Template
<img width="1196" height="672" alt="image" src="https://github.com/user-attachments/assets/71054389-d80d-4215-9c57-dd9ffa46e878" />


## License
MIT
