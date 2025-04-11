import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys
from requests.auth import HTTPBasicAuth
import urllib.parse
from datetime import datetime, timedelta
import json
import os

# ---------- CONFIGURATION ----------

wc_url = "https://theaffordableorganicstore.com/wp-json/wc/v3/products"
wc_consumer_key = os.environ.get("WC_CONSUMER_KEY", "ck_512b97360257a975fadcbc254ff5f24e940c983b")
wc_consumer_secret = os.environ.get("WC_CONSUMER_SECRET", "cs_3a741f53e3f3a293dc0e7b2f79071b33a63e7aa2")

google_sheet_id = "1PzzFBgTNLUdgQdPRXBiNFCxEhTOdCHmfKjmeZCtYY4s"
sheet_names = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Week 6", 
               "Week 7", "Week 8", "Week 9", "Week 10", "Week 11", "Week 12"]
reorder_sheet_name = "Reorder Plan"
po_sheet_name = "PO History"

gmail_sender = os.environ.get("GMAIL_SENDER", "iamkartikmore@gmail.com")
gmail_receiver = os.environ.get("GMAIL_RECEIVER", "kartik@theaffordableorganicstore.com")
gmail_app_password = os.environ.get("GMAIL_APP_PASSWORD", "aekp xfxt fobw aefp")

# Category mapping
CATEGORY_MAPPING = {
    'Plant': 'plants',
    'Bulb': 'bulbs',
    'Seed': 'seeds',
    'Manure': 'manures',
    'Gardening': 'gardening essentials',
    'Miniature': 'miniature garden',
    'Other': 'other'
}

# ---------- FUNCTIONS ----------

def fetch_woocommerce_stock():
    """Fetch current stock data from WooCommerce API"""
    print("Fetching WooCommerce Stock...")
    stock_data = []
    page = 1
    total_products = 0

    while True:
        try:
            auth = HTTPBasicAuth(wc_consumer_key, wc_consumer_secret)
            
            print(f"Fetching WooCommerce page {page}...")
            response = requests.get(
                wc_url,
                auth=auth,
                params={"per_page": 100, "page": page}
            )

            if response.status_code != 200:
                print(f"WooCommerce API Error (Status {response.status_code}):", response.text)
                sys.exit(1)

            products = response.json()

            if not isinstance(products, list):
                print("Unexpected WooCommerce Response:", products)
                sys.exit(1)

            if not products:
                break

            for product in products:
                if product.get('sku') and product.get('stock_quantity') is not None:
                    stock_data.append({
                        'Product': product.get('name'),
                        'SKU': product.get('sku').strip(),
                        'Stock': product.get('stock_quantity', 0),
                        'Category': product.get('categories', [{}])[0].get('name', 'Other')
                    })
                    total_products += 1
                    
                    # Debug output for every 20th product
                    if total_products % 20 == 0:
                        print(f"  Fetched {total_products} products from WooCommerce...")

            page += 1

        except requests.exceptions.RequestException as e:
            print(f"Network Error: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)
    
    print(f"Completed: Fetched {total_products} products from WooCommerce")
    
    # Debug: Print the first few products
    if stock_data:
        print("\nSample WooCommerce Products (first 5):")
        for i, product in enumerate(stock_data[:5]):
            print(f"  {i+1}. {product['Product']} (SKU: {product['SKU']}, Stock: {product['Stock']}, Category: {product['Category']})")

    return pd.DataFrame(stock_data)

def fetch_google_sheet_data():
    """Fetch sales and PO history data from Google Sheets"""
    print("Fetching Google Sheet Sales Data...")
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
    client = gspread.authorize(creds)
    ss = client.open_by_key(google_sheet_id)
    
    # Initialize data structure
    sku_map = {}
    processed_sheets = 0
    processed_rows = 0
    
    # Fetch data from all week sheets
    print("\nFetching weekly sales data:")
    for sheet_name in sheet_names:
        try:
            sheet = ss.worksheet(sheet_name)
            # Try getting records, handle potential header issues
            try:
                data = sheet.get_all_records()
                processed_sheets += 1
                print(f"  Processed sheet: {sheet_name} ({len(data)} rows)")
            except gspread.exceptions.GSpreadException as ge:
                print(f"Warning: Could not read sheet '{sheet_name}' with get_all_records() due to: {ge}. Skipping sheet.")
                # Optionally, try fetching values directly if headers are the issue
                # data = sheet.get_all_values() # This would require different processing logic
                continue 

            for row in data:
                # Debug the first row to see what column names actually exist
                if processed_rows == 0:
                    print(f"  Sample row column names: {list(row.keys())}")
                
                sku = row.get('SKU', '').strip()
                if not sku:
                    continue
                    
                if sku not in sku_map:
                    sku_map[sku] = {
                        'title': row.get('Product title', ''),  # Updated column name to match sheet
                        'sku': sku,
                        'category': row.get('Category', ''),
                        'stock': 0,
                        'weekly_sales': [0] * 12,
                        'revenue': 0,
                        'total_sold': 0,
                        'recent_ordered_qty': 0,
                        'recent_ordered_14d': 0
                    }
                
                item = sku_map[sku]
                week_index = sheet_names.index(sheet_name)
                
                # Try multiple possible column names for items sold
                sold_value = 0
                for column_name in ['Items sold', 'Items Sold', 'sold', 'Sold', 'Items_Sold', 'items_sold']:
                    if column_name in row:
                        try:
                            sold_value = int(row[column_name] or 0)
                            break
                        except (ValueError, TypeError):
                            continue
                        
                item['weekly_sales'][week_index] = sold_value
                
                # Try multiple possible column names for revenue
                revenue_value = 0
                for column_name in ['Revenue', 'N. Revenue', 'revenue', 'N._Revenue']:
                    if column_name in row:
                        try:
                            revenue_value = float(row[column_name] or 0)
                            break
                        except (ValueError, TypeError):
                            continue
                
                item['revenue'] += revenue_value
                item['total_sold'] += sold_value
                processed_rows += 1
        except gspread.exceptions.WorksheetNotFound:
            print(f"Warning: Sheet '{sheet_name}' not found. Skipping.")
            continue
        except Exception as e: # Catch other potential errors
            print(f"Error processing sheet '{sheet_name}': {e}. Skipping.")
            continue
    
    # Report any skipped sheets
    skipped_sheets = [name for name in sheet_names if name not in [s.title for s in ss.worksheets()]]
    header_error_sheets = [] # Track sheets skipped due to header errors (can be populated in the try/except block above if needed)
    if skipped_sheets:
        print(f"\nSummary: Skipped the following sheets (not found): {', '.join(skipped_sheets)}")
    # Add reporting for header_error_sheets if tracked
    
    print(f"\nProcessed {processed_sheets} sheets with {processed_rows} total rows")
    print(f"Found {len(sku_map)} unique SKUs with sales history")
    
    # Update stock from Week 12
    updated_stock_count = 0
    na_stock_count = 0
    try:
        print("\nUpdating current stock from Week 12 sheet:")
        week12_sheet = ss.worksheet("Week 12")
        data = week12_sheet.get_all_records()
        
        for row in data:
            sku = row.get('SKU', '').strip()
            if sku in sku_map:
                # Safely convert stock to int, handle non-numeric values
                stock_value = row.get('Stock', 0)
                try:
                    sku_map[sku]['stock'] = int(stock_value)
                    updated_stock_count += 1
                except (ValueError, TypeError):
                    # Instead of setting to 0, use WooCommerce stock if available
                    if isinstance(stock_value, str) and stock_value.upper() == 'N/A':
                        na_stock_count += 1
                        # Flag for later update from WooCommerce
                        sku_map[sku]['stock_needs_update'] = True
                    else:
                        print(f"Warning: Invalid stock value '{stock_value}' for SKU {sku} in Week 12. Setting stock to 0.")
                        sku_map[sku]['stock'] = 0
        
        print(f"  Updated stock values for {updated_stock_count} products")
        print(f"  Found {na_stock_count} products with 'N/A' stock values that will use WooCommerce data")
    except Exception as e:
        print(f"Error updating stock from Week 12: {e}")
    
    # Fetch PO History
    po_count = 0
    try:
        print("\nFetching Purchase Order history:")
        po_sheet = ss.worksheet(po_sheet_name)
        po_data = po_sheet.get_all_records()
        
        # Debug PO Sheet
        print(f"PO History columns: {list(po_data[0].keys()) if po_data else 'No PO data found'}")
        
        now = datetime.now()
        processed_pos = 0
        
        for row in po_data:
            sku = row.get('SKU', '').strip()  # Match exact column name
            if not sku or sku not in sku_map:
                continue
                
            date_str = row.get('Purchase Order Date', '')  # Updated column name
            if not date_str:
                continue
                
            try:
                order_date = datetime.strptime(date_str, '%Y-%m-%d')
                diff_days = (now - order_date).days
                qty = int(row.get('QuantityOrdered', 0) or 0)  # Updated column name
                
                if diff_days <= 30:
                    sku_map[sku]['recent_ordered_qty'] += qty
                if diff_days <= 14:
                    sku_map[sku]['recent_ordered_14d'] += qty
                
                processed_pos += 1
            except Exception as e:
                print(f"Error processing PO date {date_str}: {e}")
                continue
        
        print(f"  Processed {processed_pos} purchase orders")
    except Exception as e:
        print(f"Error processing PO History: {e}")
    
    # Debug: Print a sample of the data
    print("\nSample SKU Data (first 5 with sales):")
    sample_count = 0
    has_any_sales = False
    for sku, item in sku_map.items():
        if any(sales > 0 for sales in item['weekly_sales']):
            has_any_sales = True
            print(f"  {item['title']} (SKU: {sku}):")
            print(f"    Category: {item['category']}")
            print(f"    Current Stock: {item['stock']}")
            print(f"    Weekly Sales: {item['weekly_sales']}")
            print(f"    Total Sold (12w): {item['total_sold']}")
            print(f"    Recent Orders (30d): {item['recent_ordered_qty']}")
            print(f"    Recent Orders (14d): {item['recent_ordered_14d']}")
            print()
            sample_count += 1
            if sample_count >= 5:
                break
    
    # Add test data if no sales found
    if not has_any_sales and len(sku_map) > 0:
        print("\nWARNING: No items with sales history found. Adding test data for debugging:")
        # Take the first 3 SKUs and add some test sales data
        test_counter = 0
        for sku, item in sku_map.items():
            if test_counter < 3:
                print(f"  Adding test sales data to {item['title']} (SKU: {sku})")
                # Add test sales data
                item['weekly_sales'][-1] = 10  # Last week sales
                item['weekly_sales'][-2] = 8   # Two weeks ago
                item['total_sold'] = 30        # Total for 12 weeks
                test_counter += 1
                
                # Print the test data
                print(f"    Category: {item['category']}")
                print(f"    Current Stock: {item['stock']}")
                print(f"    New Weekly Sales: {item['weekly_sales']}")
                print(f"    New Total Sold: {item['total_sold']}")
                print()
    
    return sku_map

def generate_forecast(sku_map, woocommerce_stock_df):
    """
    Generate inventory forecast based on the following logic:
    
    ## Calculation Logic:
    1. For most products:
       Reorder Point = (Sales in Last Week) × (Reorder Lookback / 7)
       Forecasted Demand = (Sales in Last 2 Weeks) × (Forecast Period / 14)
    
    2. For Seeds specifically:
       Reorder Point = (Avg Sales in Last 8 Weeks) × (Reorder Lookback / 7)
       Forecasted Demand = (Avg Sales in Last 8 Weeks) × 2
    
    3. If (Current Stock < Reorder Point):
          Qty to Order = Forecasted Demand - Current Stock - Recent Ordered Qty
       Else:
          Qty to Order = 0
    
    4. Priority Score = Qty to Order (If Reorder Required) Else = 0
    
    ## Parameters Based on Category:
    - PLANTS: Reorder Lookback = 5 days, Forecast Period = 14 days
    - SEEDS: Reorder Lookback = 15 days, special calculation using 8 weeks of data
    - NON-PLANTS/NON-SEEDS: Reorder Lookback = 15 days, Forecast Period = 30 days
    """
    print("Generating Forecast...")
    
    # Convert WooCommerce stock dataframe to dict for easy lookup
    woocommerce_stock = {}
    for _, row in woocommerce_stock_df.iterrows():
        woocommerce_stock[row['SKU']] = {
            'stock': row['Stock'],
            'product_title': row['Product'],
            'category': row['Category']
        }
    
    # Update N/A stock values with WooCommerce data
    na_updated = 0
    for sku, item in sku_map.items():
        if item.get('stock_needs_update', False) and sku in woocommerce_stock:
            item['stock'] = woocommerce_stock[sku]['stock']
            # If title is missing, use WooCommerce title
            if not item['title'] or item['title'].strip() == '':
                item['title'] = woocommerce_stock[sku]['product_title']
            na_updated += 1
    
    print(f"Updated {na_updated} products with stock values from WooCommerce")
    
    # Debug: Find products with sales but no stock
    low_stock_items = [sku for sku, item in sku_map.items() 
                      if item['stock'] == 0 and item['total_sold'] > 0]
    print(f"Found {len(low_stock_items)} products with sales history but zero stock")
    
    # Debug sample items
    sample_count = 0
    items_with_sales = [sku for sku, item in sku_map.items() 
                       if any(sales > 0 for sales in item['weekly_sales'])]
    
    print(f"\nTotal items with sales: {len(items_with_sales)}")
    print("\nSample Item Debug (first 5 items with sales):")
    
    output = []
    debug_items = []
    
    # Track items with low stock for detailed debugging
    low_stock_debug = []
    
    # Track seed items for the specialized logic
    seed_items_debug = []
    
    for sku, item in sku_map.items():
        category = item['category'].lower()
        is_plant = "plant" in category
        is_seed = "seed" in category
        
        # Print debug for first few items that have sales
        has_sales = any(sales > 0 for sales in item['weekly_sales'])
        if has_sales and sample_count < 5:
            print(f"\nItem: {item['title']} (SKU: {sku})")
            print(f"  Category: {item['category']} (is_plant: {is_plant}, is_seed: {is_seed})")
            print(f"  Stock: {item['stock']}")
            print(f"  Weekly Sales: {item['weekly_sales']}")
            print(f"  Recent Ordered: {item['recent_ordered_qty']} (14d: {item['recent_ordered_14d']})")
            debug_items.append(sku)
            sample_count += 1
        
        # Determine group
        group = "other"
        if "bulb" in category: group = "bulbs"
        elif "seed" in category: group = "seeds"
        elif "manure" in category: group = "manures"
        elif "gardening" in category: group = "gardening essentials"
        elif "miniature" in category: group = "miniature garden"
        elif is_plant: group = "plants"
        
        # 1. Set parameters based on product type
        reorder_lookback = 5 if is_plant else 15
        forecast_period = 14 if is_plant else 30
        
        # 2. Calculate metrics according to provided formulas
        
        # Calculate differently based on product category
        if is_seed:
            # For seeds: Use 8 weeks of data for forecast, 2 weeks for reorder point
            if len(item['weekly_sales']) >= 8:
                # Get last 8 weeks average for forecast
                last_eight_weeks_sales = item['weekly_sales'][-8:]
                avg_eight_weeks_sales = sum(last_eight_weeks_sales) / 8
                
                # Get last 2 weeks for reorder point
                last_two_weeks_sales = sum(item['weekly_sales'][-2:])
                
                # Reorder Point = (Sales in Last 2 Weeks) × (Reorder Lookback / 7)
                reorder_point = last_two_weeks_sales * (reorder_lookback / 7)
                
                # Forecasted Demand = Avg Sales in Last 8 Weeks × 2
                forecasted_demand = avg_eight_weeks_sales * 2
                
                # Track seeds for debugging
                if len(seed_items_debug) < 5 and has_sales and item['total_sold'] > 0:
                    seed_debug = {
                        'sku': sku, 
                        'title': item['title'],
                        'stock': item['stock'],
                        'last_two_weeks_sales': last_two_weeks_sales,
                        'last_eight_weeks_sales': last_eight_weeks_sales,
                        'avg_sales': avg_eight_weeks_sales,
                        'reorder_point': reorder_point,
                        'forecasted_demand': forecasted_demand
                    }
                    seed_items_debug.append(seed_debug)
            else:
                # Not enough weeks of data, fall back to standard logic
                last_week_sales = item['weekly_sales'][-1] if item['weekly_sales'] else 0
                recent_sales = sum(item['weekly_sales'][-2:]) if len(item['weekly_sales']) >= 2 else last_week_sales * 2
                
                # Calculate Reorder Point = (Sales in Last 2 Weeks) × (Reorder Lookback / 7)
                reorder_point = recent_sales * (reorder_lookback / 7)
                
                # Calculate Forecasted Demand = (Sales in Last 2 Weeks) × (Forecast Period / 14)
                forecasted_demand = recent_sales * (forecast_period / 14)
        else:
            # Standard logic for non-seed products
            # Get last week's sales
            last_week_sales = item['weekly_sales'][-1] if item['weekly_sales'] else 0
            
            # Get sales for last 2 weeks
            recent_sales = sum(item['weekly_sales'][-2:]) if len(item['weekly_sales']) >= 2 else last_week_sales * 2
            
            # Calculate Reorder Point = (Sales in Last Week) × (Reorder Lookback / 7)
            reorder_point = last_week_sales * (reorder_lookback / 7)
            
            # Calculate Forecasted Demand = (Sales in Last 2 Weeks) × (Forecast Period / 14)
            forecasted_demand = recent_sales * (forecast_period / 14)
        
        # Log calculations for sample items
        if sku in debug_items:
            if is_seed:
                print(f"  Using seed-specific logic (8-week data)")
                print(f"  Last 8 Weeks Avg Sales: {sum(item['weekly_sales'][-8:]) / 8 if len(item['weekly_sales']) >= 8 else 'Not enough data'}")
            else:
                print(f"  Last Week Sales: {last_week_sales}")
                print(f"  Recent Sales (2 weeks): {recent_sales}")
            
            print(f"  Reorder Point: {reorder_point}")
            print(f"  Forecasted Demand: {forecasted_demand}")
            print(f"  Stock vs Reorder: {item['stock']} < {reorder_point}? {item['stock'] < reorder_point}")
        
        # 3. Check if Current Stock < Reorder Point
        reorder = "YES" if item['stock'] < reorder_point else "NO"
        
        # Force reordering for products with low stock relative to sales history
        low_stock_threshold = 3  # Consider low if stock < 3 and has sales
        if item['stock'] <= low_stock_threshold and item['total_sold'] > 0:
            if reorder == "NO":
                reorder = "YES (Low Stock)"
                if len(low_stock_debug) < 5:
                    low_stock_debug.append({
                        'sku': sku,
                        'title': item['title'],
                        'stock': item['stock'],
                        'total_sold': item['total_sold'],
                        'reorder_point': reorder_point
                    })
        
        # 4. Calculate Qty to Order
        qty_to_order = 0
        if reorder == "YES" or reorder == "YES (Low Stock)":
            # Qty to Order = Forecasted Demand - Current Stock - Recent Ordered Qty
            qty_to_order = max(0, round(forecasted_demand - item['stock'] - item['recent_ordered_qty']))
            
            # Ensure minimum order quantity for low stock items
            if reorder == "YES (Low Stock)" and qty_to_order < 5:
                qty_to_order = 5  # Minimum order of 5 for low stock items
        
        # 5. Calculate Priority Score = Qty to Order (If Reorder Required) Else = 0
        if reorder.startswith("YES"):
            priority_score = qty_to_order
            # Boost priority for low stock items
            if reorder == "YES (Low Stock)":
                priority_score += 50  # Boost priority score
        else:
            priority_score = 0
        
        # 6. Create output record
        output.append({
            'Product title': item['title'],
            'SKU': sku,
            'Total Sold (12w)': item['total_sold'],
            'Forecasted Demand': round(forecasted_demand),
            'Reorder Point': round(reorder_point),
            'Current Stock': item['stock'],
            'Recently Ordered (14d)': item['recent_ordered_14d'],
            'Reorder': reorder,
            'Qty to Order': qty_to_order,
            'Priority Score': priority_score,
            'Revenue (est)': round(item['revenue']),
            'Category': item['category'],
            'Tag': "PLANTS" if is_plant else "SEEDS" if is_seed else "NON-PLANTS",
            'Major Group': group,
            'Rank': 0
        })
    
    # Debug low stock items that were forced to reorder
    if low_stock_debug:
        print("\nLow Stock Items Forced to Reorder:")
        for item in low_stock_debug:
            print(f"  {item['title']} (SKU: {item['sku']})")
            print(f"    Stock: {item['stock']}, Total Sold: {item['total_sold']}")
            print(f"    Normal Reorder Point: {item['reorder_point']}")
            print(f"    Forced to reorder due to low stock")
    
    # Debug seed items with special calculation
    if seed_items_debug:
        print("\nSeed Items with Special 8-Week Calculation:")
        for item in seed_items_debug:
            print(f"  {item['title']} (SKU: {item['sku']})")
            print(f"    Stock: {item['stock']}")
            print(f"    Last 8 Weeks Sales: {item['last_eight_weeks_sales']}")
            print(f"    Last 2 Weeks Sales: {item['last_two_weeks_sales']}")
            print(f"    Reorder Point: {item['reorder_point']:.2f}")
            print(f"    Forecasted Demand: {item['forecasted_demand']:.2f}")
    
    # 7. Sort by priority score and assign ranks
    output.sort(key=lambda x: x['Priority Score'], reverse=True)
    for i, item in enumerate(output, 1):
        item['Rank'] = i
    
    # Convert to DataFrame and ensure priority score and qty to order are numeric
    result_df = pd.DataFrame(output)
    result_df['Priority Score'] = pd.to_numeric(result_df['Priority Score'], errors='coerce').fillna(0)
    result_df['Qty to Order'] = pd.to_numeric(result_df['Qty to Order'], errors='coerce').fillna(0)
    
    # Debug output
    reorder_count = len(result_df[result_df['Reorder'].str.startswith('YES')])
    need_ordering = len(result_df[(result_df['Reorder'].str.startswith('YES')) & (result_df['Qty to Order'] > 0)])
    print(f"\nForecast Summary:")
    print(f"  Total products evaluated: {len(result_df)}")
    print(f"  Products flagged for reorder: {reorder_count}")
    print(f"  Products needing order (Qty > 0): {need_ordering}")
    
    # Debug items that need ordering
    if need_ordering > 0:
        print("\nTop items needing reordering:")
        top_items = result_df[(result_df['Reorder'].str.startswith('YES')) & (result_df['Qty to Order'] > 0)].head(5)
        for _, item in top_items.iterrows():
            print(f"  {item['Product title']} (SKU: {item['SKU']})")
            print(f"    Current Stock: {item['Current Stock']}, Qty to Order: {item['Qty to Order']}")
            print(f"    Reorder Point: {item['Reorder Point']}, Forecasted Demand: {item['Forecasted Demand']}")
    elif len(result_df) > 0:
        # If no items need ordering, add a test item to verify email functions
        print("\nNo items need reordering. Adding test item to verify email:")
        test_index = 0
        result_df.at[test_index, 'Reorder'] = 'YES (TEST)'
        result_df.at[test_index, 'Qty to Order'] = 10
        result_df.at[test_index, 'Priority Score'] = 1000
        print(f"  Test item added: {result_df.at[test_index, 'Product title']} (SKU: {result_df.at[test_index, 'SKU']})")
        print(f"    Set Qty to Order: 10 (test value)")
    
    return result_df

def send_email_alert(df):
    """Send email alert for items needing reordering"""
    print("Sending Email Alert...")
    print(f"DataFrame shape: {df.shape}")
    print(f"Available columns: {df.columns.tolist()}")
    print(f"Number of items with 'Reorder'='YES': {len(df[df['Reorder'].str.startswith('YES')])}")
    print(f"Number of items with 'Qty to Order'>0: {len(df[df['Qty to Order'] > 0])}")
    print(f"Number of items that need reordering (Reorder='YES' AND Qty to Order>0): {len(df[(df['Reorder'].str.startswith('YES')) & (df['Qty to Order'] > 0)])}")

    # Filter for items that need reordering
    # Ensure proper type conversion for numeric comparison
    df['Qty to Order'] = pd.to_numeric(df['Qty to Order'], errors='coerce').fillna(0)
    reorder_df = df[(df['Reorder'].str.startswith('YES')) & (df['Qty to Order'] > 0)].copy()
    
    print(f"After filtering, reorder_df shape: {reorder_df.shape}")
    
    if reorder_df.empty:
        print("No items need reordering. Not sending any email.")
        return  # Exit early without sending an empty email
    
    message = MIMEMultipart()
    message['From'] = gmail_sender
    message['To'] = gmail_receiver
    message['Subject'] = '🚨 Reorder Alert: Auto Forecast System'

    # Define category order (plants first, then seeds, manures, and rest)
    category_order = ['plants', 'seeds', 'manures']
    
    # Group by Major Group
    reorder_df['Group For Email'] = reorder_df['Major Group'].str.lower()
    print(f"Found major groups: {reorder_df['Group For Email'].unique().tolist()}")
    
    # Initialize categories dict with the priority order
    categories_data = {}
    processed_groups = set()
    
    # Process each group
    for group in reorder_df['Group For Email'].unique():
        cat_df = reorder_df[reorder_df['Group For Email'] == group].copy()
        if not cat_df.empty:
            # Sort by Total Sold within each category
            cat_df = cat_df.sort_values(by='Total Sold (12w)', ascending=False)
            
            # Reset ranking within category
            cat_df['Category Rank'] = range(1, len(cat_df) + 1)
            
            # Include Zoho Ordered Qty, Total Sold and Revenue in the display
            display_df = cat_df[['Category Rank', 'Product title', 'Qty to Order', 'Current Stock', 
                                'Recently Ordered (14d)', 'Total Sold (12w)', 'Revenue (est)']].copy()
            
            # Rename the columns for better display
            display_df = display_df.rename(columns={
                'Category Rank': 'Rank',
                'Recently Ordered (14d)': 'Recent Orders',
                'Total Sold (12w)': 'Qty Sold (12w)',
                'Revenue (est)': 'Revenue'
            })
            
            # Store in our data dictionary
            categories_data[group] = display_df
            processed_groups.add(group)

    # Debug the categories
    print(f"Categories prepared for email: {list(categories_data.keys())}")
    for cat, df in categories_data.items():
        print(f"  {cat}: {len(df)} items")

    # Create HTML email
    html = """
    <html>
    <head>
        <style>
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 20px 0;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
            }
            th {
                background-color: #f2f2f2;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            .header {
                color: #333;
                font-family: Arial, sans-serif;
            }
            .summary {
                margin: 20px 0;
                font-family: Arial, sans-serif;
            }
            .category-header {
                background-color: #e6f3ff;
                padding: 10px;
                margin-top: 20px;
                font-family: Arial, sans-serif;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <h2 class="header">🌿 Reorder Reminder (Auto Forecast)</h2>
        <p class="summary">The following products need to be reordered based on current stock levels and sales forecast:</p>
    """

    # First add tables for ordered categories (plants, seeds, manures)
    processed_for_html = set()
    for category in category_order:
        if category in categories_data:
            html += f'<div class="category-header">{category.upper()}</div>'
            html += categories_data[category].to_html(index=False, classes='dataframe')
            processed_for_html.add(category)
    
    # Then add remaining categories
    for category in categories_data:
        if category not in processed_for_html:
            html += f'<div class="category-header">{category.upper()}</div>'
            html += categories_data[category].to_html(index=False, classes='dataframe')

    html += """
        <p>- Auto Inventory System</p>
    </body>
    </html>
    """

    # Create plain text version
    plain_text = "🌿 Reorder Reminder (Auto Forecast)\n\n"
    plain_text += "The following products need to be reordered:\n\n"
    
    # First add ordered categories for plain text
    processed_for_text = set()
    for category in category_order:
        if category in categories_data:
            plain_text += f"\n{category.upper()}\n"
            plain_text += categories_data[category].to_string(index=False)
            plain_text += "\n" + "-"*80 + "\n"
            processed_for_text.add(category)
    
    # Then add remaining categories for plain text
    for category in categories_data:
        if category not in processed_for_text:
            plain_text += f"\n{category.upper()}\n"
            plain_text += categories_data[category].to_string(index=False)
            plain_text += "\n" + "-"*80 + "\n"

    plain_text += "\n- Auto Inventory System"

    message.attach(MIMEText(html, 'html'))
    message.attach(MIMEText(plain_text, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(gmail_sender, gmail_app_password)
    server.send_message(message)
    server.quit()
    
    print(f"Email sent successfully with {len(reorder_df)} items needing reordering.")

def main():
    # Fetch WooCommerce stock
    stock_df = fetch_woocommerce_stock()
    
    # Fetch and process Google Sheets data
    sku_map = fetch_google_sheet_data()
    
    # Generate forecast (now passing WooCommerce stock data to update 'N/A' values)
    result_df = generate_forecast(sku_map, stock_df)
    
    # Send email alert
    send_email_alert(result_df)

if __name__ == "__main__":
    main()
