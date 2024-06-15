import yfinance as yf
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
import pandas as pd

# Define the tickers for various stocks
tickers = {
    "ENEL": "ENEL.MI",
    "INTESA SANPAOLO": "ISP.MI",
    "POSTE ITALIANE": "PST.MI",
    "BANCO BPM": "BAMI.MI",
    "STELLANTIS": "STLAM.MI",  # Verifica che questo ticker sia corretto
    "GENERALI": "G.MI"
}

# Create a dictionary to store prices
prices = {}

# Get the closing prices of the stocks
for company, ticker in tickers.items():
    stock = yf.Ticker(ticker)
    hist = stock.history(period="1d")
    if not hist.empty:
        price = hist['Close'].iloc[-1]
        prices[company] = price
    else:
        print(f"No data found for {company}")

# Path to the Excel file
file_path = r'.....'  # Replace with the correct path to your Excel file

try:
    # Try to open the existing Excel file
    workbook = load_workbook(file_path)
    print("Opened existing Excel file.")
except:
    # If the file is not a valid Excel file, create a new one
    workbook = Workbook()
    print("Created a new Excel file.")
    workbook.save(file_path)

sheet = workbook.active

# Update the sheet with stock prices
if sheet.title == 'Sheet':
    sheet.title = 'Stock Data'
for row in range(1, sheet.max_row + 1):
    cell_value = sheet.cell(row=row, column=4).value
    if cell_value in prices:
        sheet.cell(row=row, column=6, value=prices[cell_value])  # Update 'Valore per azione'

# Create a new sheet for the charts
chart_sheet = workbook.create_sheet(title="Stock Charts")

# Generate and add charts to the new sheet
for idx, (company, ticker) in enumerate(tickers.items(), start=1):
    stock = yf.Ticker(ticker)
    hist = stock.history(period="6mo")  # Get data for the last 6 months
    if not hist.empty:
        # Remove timezone info from datetime
        hist.index = hist.index.tz_localize(None)
        
        # Add the data to the new sheet
        chart_sheet.append([company])
        for date, close in zip(hist.index, hist['Close']):
            chart_sheet.append([date, close])
        
        # Create a chart
        chart = LineChart()
        chart.title = f'{company} Stock Price'
        chart.y_axis.title = 'Price'
        chart.x_axis.title = 'Date'

        data = Reference(chart_sheet, min_col=2, min_row=idx*len(hist)+1, max_row=(idx+1)*len(hist), max_col=2)
        dates = Reference(chart_sheet, min_col=1, min_row=idx*len(hist)+1, max_row=(idx+1)*len(hist))
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(dates)

        # Position the chart in the sheet
        chart.anchor = f'A{idx * 15}'
        chart_sheet.add_chart(chart)
    else:
        print(f"No historical data found for {company}")

# Save the Excel file
workbook.save(file_path)

print("Updated stock prices and generated charts in Excel:")
for company, price in prices.items():
    print(f"{company}: {price}")
