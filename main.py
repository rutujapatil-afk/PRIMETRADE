import requests
import pandas as pd
import time
from openpyxl import load_workbook
from datetime import datetime
import os

def fetch_crypto_data():
    """Fetches live cryptocurrency data from CoinGecko API with retries."""
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": "false"
    }
    for _ in range(3):  # Retry up to 3 times
        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            print(f"⚠️ Error fetching data: {e}. Retrying...")
            time.sleep(5)  # Wait 5 seconds before retrying
    print("❌ Failed to fetch data after 3 attempts.")
    return []

def process_data(data):
    """Processes raw API data into a structured DataFrame."""
    return pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])

def update_excel(df, file_name="crypto_data.xlsx"):
    """Updates Excel file with new cryptocurrency data."""
    try:
        if os.path.exists(file_name):
            with pd.ExcelWriter(file_name, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name="Live Data", index=False)
        else:
            with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Live Data", index=False)
    except PermissionError:
        print("⚠️ Excel file is open! Close it to update data.")
    except Exception as e:
        print(f"❌ Error updating Excel: {e}")
    print(f"✅ Excel updated at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

def analyze_data(df):
    """Performs analysis on the cryptocurrency data and saves a report."""
    top_5 = df.nlargest(5, "market_cap")
    avg_price = df["current_price"].mean()
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax()]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin()]

    report = f"""
    Cryptocurrency Data Analysis Report
    -----------------------------------
    Top 5 Cryptocurrencies by Market Cap:
    {top_5[['name', 'market_cap']].to_string(index=False)}
    
    Average Price of Top 50 Cryptocurrencies: ${avg_price:.2f}
    
    Highest 24h Change: {highest_change['name']} ({highest_change['price_change_percentage_24h']:.2f}%)
    Lowest 24h Change: {lowest_change['name']} ({lowest_change['price_change_percentage_24h']:.2f}%)
    """
    
    with open("crypto_analysis.txt", "w", encoding="utf-8") as f:
        f.write(report)
    
    print(report)

def main():
    """Main loop that fetches and updates crypto data every 5 minutes."""
    while True:
        print("\n⏳ Fetching latest cryptocurrency data...")
        data = fetch_crypto_data()
        if data:
            df = process_data(data)
            update_excel(df)
            analyze_data(df)
        else:
            print("❌ Skipping update due to data fetch failure.")
        print("🔄 Waiting for 5 minutes before next update...")
        time.sleep(300)

if __name__ == "__main__":
    main()
