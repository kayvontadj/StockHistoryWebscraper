import fix_yahoo_finance as yf 
import os
import xlrd
import pandas as pd
import csv
from pandas_datareader import data as pdr

#fixes yahoo finance issues
yf.pdr_override()

path_closing_prices = os.environ["HOMEPATH"] + "\\Desktop\\Closing Prices"

#makes folder if one doesn't exist
if not os.path.isdir(path_closing_prices):
	os.makedirs(path_closing_prices)
	print("Making 'Closing Prices' Folder in Desktop")


#get tickers
workbook = xlrd.open_workbook("StockList.xlsx")
sheet = workbook.sheet_by_index(0)
tickers_shorts = [sheet.cell_value(r,2) for r in range(len(sheet.col_values(2)))]
tickers_longs = [sheet.cell_value(r,0) for r in range(len(sheet.col_values(0)))]

path_shorts = path_closing_prices + "\\" + tickers_shorts[0] + ".csv"
path_longs = path_closing_prices + "\\" + tickers_longs[0] + ".csv"

historical_data_shorts = pdr.get_data_yahoo(tickers_shorts[0], start="2013-09-30", end="2018-09-30")
historical_data_csv_shorts = historical_data_shorts.to_csv()
historical_data_longs = pdr.get_data_yahoo(tickers_longs[0], start="2013-09-30", end="2018-09-30")
historical_data_csv_longs = historical_data_longs.to_csv()

csv_shorts = open(path_shorts, "w")
csv_shorts.write(historical_data_csv_shorts)
csv_shorts.close()

csv_longs = open(path_longs, "w")
csv_longs.write(historical_data_csv_longs)
csv_longs.close()


data_date_shorts = pd.read_csv(path_shorts)["Date"].tolist()
data_date_longs = pd.read_csv(path_longs)["Date"].tolist()

with open(path_closing_prices + "\\Temp_Shorts.csv", 'w', newline='') as f:
	thewriter = csv.writer(f, dialect='excel')
	thewriter.writerow(["Date"] + data_date_shorts)
	for ticker in tickers_shorts:
		try:
			path = path_closing_prices + "\\" + ticker + ".csv"

			historical_data = pdr.get_data_yahoo(ticker, start="2013-09-30", end="2018-09-30")
			historical_data_csv = historical_data.to_csv()

			csv_shorts = open(path, "w")
			csv_shorts.write(historical_data_csv)
			csv_shorts.close()

			data_close = pd.read_csv(path)["Close"].tolist()

			thewriter.writerow([ticker] + data_close)
		except ValueError:
			pass
temp = pd.read_csv(path_closing_prices + "\\Temp_Shorts.csv", index_col=0,error_bad_lines=False)
temp_transpose = temp.transpose().to_csv()
csv_transpose = open(path_closing_prices + "\\Short Closing Prices.csv", "w")
csv_transpose.write(temp_transpose)
csv_transpose.close()
with open(path_closing_prices + "\\Temp_Longs.csv", 'w', newline='') as f:
	thewriter = csv.writer(f, dialect='excel')
	thewriter.writerow(["Date"] + data_date_longs)
	for ticker in tickers_longs:
		try:
			path = path_closing_prices + "\\" + ticker + ".csv"

			historical_data = pdr.get_data_yahoo(ticker, start="2013-09-30", end="2018-09-30")
			historical_data_csv = historical_data.to_csv()

			csv_longs = open(path, "w")
			csv_longs.write(historical_data_csv)
			csv_longs.close()

			data_close = pd.read_csv(path)["Close"].tolist()

			thewriter.writerow([ticker] + data_close)
		except ValueError:
			pass
temp = pd.read_csv(path_closing_prices + "\\Temp_Longs.csv", index_col=0, error_bad_lines=False)
temp_transpose = temp.transpose().to_csv()
csv_transpose = open(path_closing_prices + "\\Long Closing Prices.csv", "w")
csv_transpose.write(temp_transpose)
csv_transpose.close()