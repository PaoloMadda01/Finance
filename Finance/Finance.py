import yfinance as yf
from openpyxl import load_workbook
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image

# Define the tickers for various stocks
tickers = {
    "ENEL": "ENEL.MI",
    "INTESA SANPAOLO": "ISP.MI",
    "POSTE ITALIANE": "PST.MI",
    "BANCO BPM": "BAMI.MI",
    "STELLANTIS": "STLA.MI",
    "GENERALI": "G.MI"
}

# Create a dictionary to store prices
prices = {}

# Get the closing prices of the stocks
for company, ticker in tickers.items():
    stock = yf.Ticker(ticker)
    price = stock.history(period="1d")['Close'].iloc[-1]
    prices[company] = price

# Path to the Excel file
file_path = '/mnt/data/your_excel_file.xlsx'  # Replace with the correct path to your Excel file

# Open the Excel file and update the values
workbook = load_workbook(file_path)
sheet = workbook.active

# Find the 'Nome' column and update the corresponding prices
for row in range(1, sheet.max_row + 1):
    cell_value = sheet.cell(row=row, column=4).value
    if cell_value in prices:
        sheet.cell(row=row, column=6, value=prices[cell_value])  # Update 'Valore per azione'

# Generate and save charts
for company, ticker in tickers.items():
    stock = yf.Ticker(ticker)
    hist = stock.history(period="6mo")  # Get data for the last 6 months
    plt.figure(figsize=(10, 5))
    plt.plot(hist.index, hist['Close'], label=f'{company}')
    plt.title(f'{company} Stock Price')
    plt.xlabel('Date')
    plt.ylabel('Price')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    chart_path = f'/mnt/data/{company}_chart.png'
    plt.savefig(chart_path)
    plt.close()
    
    # Insert the chart into the Excel file
    img = Image(chart_path)
    img.anchor = f'I{row}'  # Change the anchoring if necessary
    sheet.add_image(img)

# Save the Excel file
workbook.save(file_path)

print("Updated stock prices and generated charts in Excel:")
for company, price in prices.items():
    print(f"{company}: {price}")
