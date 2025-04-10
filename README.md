# Inventory Forecast System

Automated inventory forecasting system that integrates with WooCommerce and Google Sheets to provide daily reorder recommendations.

## Features

- Fetches current stock levels from WooCommerce
- Retrieves sales data from Google Sheets
- Calculates reorder points based on sales history
- Generates forecasts using category-specific logic
- Sends email alerts for items needing reorder
- Runs daily via GitHub Actions

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install requests pandas gspread oauth2client
   ```

3. Set up the following secrets in GitHub:
   - `GOOGLE_SERVICE_ACCOUNT`: Your Google service account JSON
   - `GMAIL_SENDER`: Sender email address
   - `GMAIL_RECEIVER`: Receiver email address
   - `GMAIL_APP_PASSWORD`: Gmail app password
   - `WC_CONSUMER_KEY`: WooCommerce consumer key
   - `WC_CONSUMER_SECRET`: WooCommerce consumer secret

4. Create a `service_account.json` file with your Google service account credentials (not included in repo)

## Usage

Run manually:
```bash
python inventory_forecast.py
```

The script will:
1. Fetch current stock from WooCommerce
2. Get sales data from Google Sheets
3. Calculate forecasts and reorder points
4. Send email alerts for items needing reorder

## Forecast Logic

### Plants
- Reorder Lookback: 5 days
- Forecast Period: 14 days

### Seeds
- Uses 8-week average for calculations
- Special reorder point calculation

### Other Products
- Reorder Lookback: 15 days
- Forecast Period: 30 days

## Email Alerts

Emails include:
- Items grouped by category
- Current stock levels
- Recommended order quantities
- Priority scores
- Recent sales data 