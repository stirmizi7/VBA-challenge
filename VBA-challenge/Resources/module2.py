import pandas as pd

# Define the paths to the Excel files
alphabetical_file_path = '/Users/miztirmizi/Desktop/VBA-challenge/Resources/alphabetical_testing.xlsx'
multiple_year_file_path = '/Users/miztirmizi/Desktop/VBA-challenge/Resources/multiple_year_stock_data.xlsx'

# Read the first Excel file into a DataFrame
alphabetical_df = pd.read_excel(alphabetical_file_path)

# Read the second Excel file into a DataFrame
multiple_year_df = pd.read_excel(multiple_year_file_path)

# Print the DataFrames (you can perform further analysis here)
print("Data from alphabetical_testing.xlsx:")
print(alphabetical_df)

print("\nData from multiple_year_stock_data.xlsx:")
print(multiple_year_df)

import pandas as pd

# Read the Excel file into a DataFrame (assuming you've already done this)
df = pd.read_excel('alphabetical_testing.xlsx')
df = pd.read_excel('Multiple_year_stock_data.xlsx')

# Create an empty DataFrame to store the results
results_df = pd.DataFrame(columns=['Ticker', 'Yearly Change', 'Percentage Change', 'Total Stock Volume'])

# Loop through each unique ticker symbol
for ticker in df['ticker'].unique():
    # Filter data for the current ticker for one year
    ticker_data = df[df['ticker'] == ticker]

    # Calculate the yearly change from opening to closing price
    yearly_change = ticker_data['close'].iloc[-1] - ticker_data['open'].iloc[0]

    # Calculate the percentage change
    opening_price = ticker_data['open'].iloc[0]
    percentage_change = (yearly_change / opening_price) * 100

    # Calculate the total stock volume
    total_stock_volume = ticker_data['vol'].sum()

    # Append the results to the results DataFrame
    results_df = results_df.append({
        'Ticker': ticker,
        'Yearly Change': yearly_change,
        'Percentage Change': percentage_change,
        'Total Stock Volume': total_stock_volume
    }, ignore_index=True)

# Print or display the results
print(results_df)

import openpyxl
import pandas as pd

# Function to calculate yearly change, percentage change, and total volume
def calculate_stock_metrics(df):
    yearly_change = df['close'].iloc[-1] - df['open'].iloc[0]
    opening_price = df['open'].iloc[0]
    percentage_change = (yearly_change / opening_price) * 100
    total_stock_volume = df['vol'].sum()
    return yearly_change, percentage_change, total_stock_volume

# List of Excel files to analyze
excel_files = ['alphabetical_testing.xlsx', 'multiple_year_stock_data.xlsx']

# Create an empty DataFrame to store the results
results_df = pd.DataFrame(columns=['Ticker', 'Yearly Change', 'Percentage Change', 'Total Stock Volume'])

# Loop through each Excel file
for excel_file in excel_files:
    # Load the Excel file using openpyxl
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    sheet = wb.active

    # Convert the sheet to a DataFrame
    df = pd.DataFrame(sheet.values, columns=[cell.value for cell in sheet[1]])

    # Calculate metrics for each unique ticker symbol
    for ticker in df['<ticker>'].unique():
        ticker_data = df[df['<ticker>'] == ticker]
        yearly_change, percentage_change, total_stock_volume = calculate_stock_metrics(ticker_data)
        
        # Append the results to the results DataFrame
        results_df = results_df.append({
            'Ticker': ticker,
            'Yearly Change': yearly_change,
            'Percentage Change': percentage_change,
            'Total Stock Volume': total_stock_volume
        }, ignore_index=True)

# Print or display the results
print(results_df)
