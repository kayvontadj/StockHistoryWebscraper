import fix_yahoo_finance as yf 
import os
import xlrd
import pandas as pd
import csv
from pandas_datareader import data as pdr

#gets ticker symbols from provided excel doc
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

#writes data file to csv
def writetoCSV(path, data):
	csv_file = open(path, "w")
	csv_file.write(data)
	csv_file.close()

#fixes yahoo finance issues
yf.pdr_override()

desired_value = input("Which value would you like to view: \nOpen	\nHigh	\nLow	\nClose	\nAdj Close \nVolume\n")
possible_values = ["open", "high", "low", "close", "adjclose", "volume"]
while desired_value.replace(' ', '').lower() not in possible_values:
	print("Choose one of the given values\n")
	desired_value = input("Which value would you like to view: \nOpen	\nHigh	\nLow	\nClose	\nAdj Close \nVolume\n")

path_values = os.environ["HOMEPATH"] + "\\Desktop\\Stock Data" 

#makes folder if one doesn't exist
if not os.path.isdir(path_values):
	os.makedirs(path_values)
	print("Making 'Stock Data' Folder in Desktop\n")

tickers = getTickers("StockList.xlsx", column = 0)

ticker_path = path_values + "\\" + tickers[0] + ".csv"

historical_data = pdr.get_data_yahoo(tickers[0], start="2013-09-30", end="2018-09-30").to_csv()

writetoCSV(ticker_path, historical_data)


dates = pd.read_csv(ticker_path)["Date"].tolist()

with open(path_values + "\\Temp.csv", 'w', newline='') as f:
	thewriter = csv.writer(f, dialect='excel')
	thewriter.writerow(["Date"] + dates)
	for ticker in tickers[0:2]:
		try:
			path = path_values + "\\" + ticker + ".csv"

			historical_data = pdr.get_data_yahoo(ticker, start="2013-09-30", end="2018-09-30").to_csv()
			csv = open(path, "w")
			csv.write(historical_data)
			csv.close()

			values = pd.read_csv(path)[desired_value.title()].tolist()

			thewriter.writerow([ticker] + values)
		except ValueError:
			pass
temp = pd.read_csv(path_values + "\\Temp.csv", index_col=0,error_bad_lines=False)
temp_transpose = temp.transpose().to_csv()
os.remove(path_values + "\\Temp.csv")
csv_transpose = open(path_values + "\\" + desired_value.title() + ".csv", "w")
csv_transpose.write(temp_transpose)
csv_transpose.close()