import fix_yahoo_finance as yf 
import os
import xlrd
import pandas as pd
import csv
from pandas_datareader import data as pdr

def getTickers(path, sheet_number=0, row=-1, column=-1):
	workbook = xlrd.open_workbook(path)
	sheet = workbook.sheet_by_index(sheet_number)
	if row == -1 and column == -1:
		print("Must specify row and/or column")
		return
	elif row != -1 and column != -1:
		tickers = [sheet.cell_value(row, column)]
	elif column != -1:
		tickers = [sheet.cell_value(r, column) for r in range(len(sheet.col_values(column)))]
	else:
		tickers = [sheet.cell_value(row, c) for c in range(len(sheet.row_values(row)))]
	return tickers

#fixes yahoo finance issues
yf.pdr_override()

path_closing_prices = os.environ["HOMEPATH"] + "\\Desktop\\Closing Prices"

#makes folder if one doesn't exist
if not os.path.isdir(path_closing_prices):
	os.makedirs(path_closing_prices)
	print("Making 'Closing Prices' Folder in Desktop")

tickers = getTickers("StockList.xlsx", column = 0)

ticker_path = path_closing_prices + "\\" + tickers[0] + ".csv"

historical_data = pdr.get_data_yahoo(tickers[0], start="2013-09-30", end="2018-09-30").to_csv()

ticker_csv = open(ticker_path, "w")
ticker_csv.write(historical_data)
ticker_csv.close()

dates = pd.read_csv(ticker_path)["Date"].tolist()

with open(path_closing_prices + "\\Temp.csv", 'w', newline='') as f:
	thewriter = csv.writer(f, dialect='excel')
	thewriter.writerow(["Date"] + dates)
	for ticker in tickers[0:2]:
		try:
			path = path_closing_prices + "\\" + ticker + ".csv"

			historical_data = pdr.get_data_yahoo(ticker, start="2013-09-30", end="2018-09-30").to_csv()
			csv = open(path, "w")
			csv.write(historical_data)
			csv.close()

			closing_prices = pd.read_csv(path)["Close"].tolist()

			thewriter.writerow([ticker] + closing_prices)
		except ValueError:
			pass
temp = pd.read_csv(path_closing_prices + "\\Temp.csv", index_col=0,error_bad_lines=False)
temp_transpose = temp.transpose().to_csv()
os.remove(path_closing_prices + "\\Temp.csv")
csv_transpose = open(path_closing_prices + "\\Closing Prices.csv", "w")
csv_transpose.write(temp_transpose)
csv_transpose.close()