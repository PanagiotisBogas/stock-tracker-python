import yfinance as yf
import openpyxl
import matplotlib.pyplot as plt


#get stock data
def get_stock_data(ticker):
    stock = yf.Ticker(ticker)
    hist = stock.history(period="5d")
    return hist


# calc stock performance
def calculate_performance(stock):
    close_prices = stock['Close']
    if len(close_prices) < 2:
        return None, None  # Not enough data to calculate performance
    last_price = close_prices.iloc[-1]
    prev_price = close_prices.iloc[-2]
    percentage_change = ((last_price - prev_price) / prev_price) * 100
    return last_price, percentage_change


#user's stocks
stocks = {
    'AAPL': 10,
    'MSFT': 5,
    'GOOGL': 2,
    'AMD': 3,
    'TSLA': 4,
    'JPM':2,
    'NFLX':6,
    'WMT': 7,
    'JNJ': 5
}

# initialize excel workbook and sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Stock Data"

#excel headers
sheet.append(["Stock", "Price $", "24h Change (%)", "Amount Owned"])

# get data and fill the excel
portfolio_worth = None
dates = None

for ticker, amount in stocks.items():
    stock_data = get_stock_data(ticker)
    last_price, percentage_change = calculate_performance(stock_data)

    if last_price is None or percentage_change is None:
        print(f"Not enough data for {ticker}")
        continue

    # append data to excel
    sheet.append([ticker, last_price, percentage_change, amount])

    # calc portfolio worth for each day and keep track of it
    daily_worth = stock_data['Close'] * amount
    if portfolio_worth is None:
        portfolio_worth = daily_worth
    else:
        portfolio_worth += daily_worth
    dates = stock_data.index

# save excel
file_name = "stocks_portfolio.xlsx"
wb.save(file_name)

# create a chart for portfolio worth over the last week
plt.figure(figsize=(10, 5))
plt.plot(dates, portfolio_worth)
plt.xlabel('Date')
plt.ylabel('Portfolio Worth')
plt.title('Portfolio Worth Over Last Week')
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()

# save plot as image
plot_image = "portfolio_worth.png"
plt.savefig(plot_image)

# add plot to excel
img = openpyxl.drawing.image.Image(plot_image)
sheet.add_image(img, 'E5')

# save excel file image
wb.save(file_name)

print(f"Excel file '{file_name}' with stock data and portfolio worth graph has been created successfully.")
