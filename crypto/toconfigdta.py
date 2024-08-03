import pandas as pd
import numpy as np

# Update the file path as needed
file_path = 'D:/cllg/web scraping/web_scraping_project/crypto/crypto_data.xlsx'

try:
    # Load the Excel file
    crypto_data = pd.read_excel(file_path)

    # Check the current columns
    print("Current columns:", crypto_data.columns)
    print("Number of columns:", len(crypto_data.columns))

    # Define the correct column names
    correct_column_names = ['Index', 'Coin', 'Symbol', 'Price', '1h Change', '24h Change', '7d Change', '30d Change', '24h Volume', 'Market Cap', 'FDV', 'Market Cap / FDV']

    # Assign the correct column names
    crypto_data.columns = correct_column_names

    # Clean and convert columns to numeric types, handle non-numeric values
    def clean_and_convert(column):
        return pd.to_numeric(crypto_data[column].replace({'\$': '', ',': '', '-': np.nan}, regex=True))

    crypto_data['Price'] = clean_and_convert('Price')
    crypto_data['24h Volume'] = clean_and_convert('24h Volume')
    crypto_data['Market Cap'] = clean_and_convert('Market Cap')
    crypto_data['FDV'] = clean_and_convert('FDV')

    # Add new calculated columns
    crypto_data['Market Cap / FDV'] = crypto_data['Market Cap'] / crypto_data['FDV']
    crypto_data['Volume to Market Cap Ratio'] = crypto_data['24h Volume'] / crypto_data['Market Cap']

    # Add Price Change Percentage columns
    crypto_data['1h Change (%)'] = pd.to_numeric(crypto_data['1h Change'].str.replace('%', ''), errors='coerce') / 100
    crypto_data['24h Change (%)'] = pd.to_numeric(crypto_data['24h Change'].str.replace('%', ''), errors='coerce') / 100
    crypto_data['7d Change (%)'] = pd.to_numeric(crypto_data['7d Change'].str.replace('%', ''), errors='coerce') / 100
    crypto_data['30d Change (%)'] = pd.to_numeric(crypto_data['30d Change'].str.replace('%', ''), errors='coerce') / 100

    # Add Volume to Market Cap Percentage column
    crypto_data['Volume to Market Cap (%)'] = crypto_data['24h Volume'] / crypto_data['Market Cap'] * 100

    # Add Market Cap Rank column
    crypto_data['Market Cap Rank'] = crypto_data['Market Cap'].rank(ascending=False)

    # Handle potential NaN values
    crypto_data.fillna(0, inplace=True)

    # Save the enhanced DataFrame back to Excel
    enhanced_file_path = 'D:/cllg/web scraping/web_scraping_project/crypto/enhanced_crypto_data.xlsx'
    crypto_data.to_excel(enhanced_file_path, index=False)

    print("Enhanced data saved to:", enhanced_file_path)

except FileNotFoundError:
    print(f"File not found: {file_path}")
except Exception as e:
    print(f"An error occurred: {e}")
