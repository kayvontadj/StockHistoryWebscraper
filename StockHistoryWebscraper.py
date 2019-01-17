import fix_yahoo_finance as yf 
import os
import xlrd
import pandas as pd
import csv
from pandas_datareader import data as pdr
from tkinter import *
import datetime as dt

values = ["Open", "High", "Low", "Close", "Adj Close", "Volume"]
desired_value = ""
start_date = ""
end_date = ""

root = Tk()

value = StringVar(root)
value.set("--Select the type of information--")

Label(root, text="Start Date").grid(row=0)
Label(root, text="End Date").grid(row=1)
start = Entry(root)
end = Entry(root)
start.grid(row=0,column=1)
end.grid(row=1,column=1)

def on_start_click(event):
    """function that gets called whenever start is clicked"""
    if start.get() == 'MM/DD/YYYY':
       start.delete(0, "end") # delete all the text in the start
       start.insert(0, '') #Insert blank for user input
       start.config(fg = 'black')
def on_end_click(event):
    """function that gets called whenever start is clicked"""
    if end.get() == 'MM/DD/YYYY':
       end.delete(0, "end") # delete all the text in the start
       end.insert(0, '') #Insert blank for user input
       end.config(fg = 'black')
def on_start_focusout(event):
    if start.get() == '':
        start.insert(0, 'MM/DD/YYYY')
        start.config(fg = 'grey')
def on_end_focusout(event):
    if end.get() == '':
        end.insert(0, 'MM/DD/YYYY')
        end.config(fg = 'grey')


start.insert(0, 'MM/DD/YYYY')
start.bind('<FocusIn>', on_start_click)
start.bind('<FocusOut>', on_start_focusout)
start.config(fg = 'grey')

end.insert(0, 'MM/DD/YYYY')
end.bind('<FocusIn>', on_end_click)
end.bind('<FocusOut>', on_end_focusout)
end.config(fg = 'grey')

value_menu = OptionMenu(root, value, *values)
value_menu.grid(row=2,column=1)

def getInputs():
	global start_date, end_date, desired_value
	desired_value = value.get()
	start_date = start.get()
	end_date = end.get()
	start_date = dt.datetime.strptime(start_date, "%m/%d/%Y").isoformat(sep='T', timespec='auto')[0:10]
	end_date = dt.datetime.strptime(end_date, "%m/%d/%Y").isoformat(sep='T', timespec='auto')[0:10]
	#start_date = start_date[6:10] + '-' + start_date[0:2] + '-' + start_date[3:5]
	#end_date = end_date[6:10] + '-' + end_date[0:2] + '-' + end_date[3:5]
	print("Chosen value is", desired_value)
	print("Chosen start date is", start_date)
	print("Chosen end date is", end_date)
	root.quit()
button = Button(root, text="Run", command=getInputs)
button.grid(row=3,column=1)
mainloop()
root.quit()
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

path_values = os.environ["HOMEPATH"] + "\\Desktop\\Stock Data" 

#makes folder if one doesn't exist
if not os.path.isdir(path_values):
	os.makedirs(path_values)
	print("Making 'Stock Data' Folder in Desktop\n")

tickers = getTickers("StockList.xlsx", column = 0)

ticker_path = path_values + "\\" + tickers[0] + ".csv"

historical_data = pdr.get_data_yahoo(tickers[0], start=start_date, end=end_date).to_csv()

writetoCSV(ticker_path, historical_data)


dates = pd.read_csv(ticker_path)["Date"].tolist()

with open(path_values + "\\Temp.csv", 'w', newline='') as f:
	thewriter = csv.writer(f, dialect='excel')
	thewriter.writerow(["Date"] + dates)
	for ticker in tickers[0:len(tickers)]:
		try:
			path = path_values + "\\" + ticker + ".csv"

			historical_data = pdr.get_data_yahoo(ticker, start=start_date, end=end_date).to_csv()
			writetoCSV(path, historical_data)

			values = pd.read_csv(path)[desired_value].tolist()

			thewriter.writerow([ticker] + values)
		except ValueError:
			pass
temp = pd.read_csv(path_values + "\\Temp.csv", index_col=0,error_bad_lines=False)
temp_transpose = temp.transpose().to_csv()
os.remove(path_values + "\\Temp.csv")
csv_transpose = open(path_values + "\\" + desired_value + ".csv", "w")
csv_transpose.write(temp_transpose)
csv_transpose.close()