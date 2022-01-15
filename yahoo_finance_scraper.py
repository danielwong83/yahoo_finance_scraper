import xlrd  # excel read
import requests  # for capturing an instance of the website being requested
import time  # for time
import xlsxwriter  # for exporting data as excel files
from datetime import datetime  # for time
from datetime import date  # for time
from time import sleep  # delay speed to not get ip banned
from bs4 import BeautifulSoup  # for parsing purposes
# from tkinter import Tk  # for buttons
# from tkinter.filedialog import askopenfilename  # to open file wanted

# runs using yahoo stocks
# syntax for yahoo stock tickers, Examples below:
# US NASDAQ: TSLA, FB, NFLX, AAPL, LULU
# HK HKSE: 0700.HK (for Tencent), 0008.HK (for PCCW), 0005.HK (for HSBC Holdings)
# UK LSE: VOD.L, HSBA.L
# KR: 005930.KS (for Samsung)

fullList = list()

def usingExcel():
    # Tk().withdraw()
    # filename = askopenfilename()
    #
    # wb = xlrd.open_workbook(filename)

    # file_name = input("Enter your file name, including .xlsx extension.")

    wb = xlrd.open_workbook("slow.xlsx")
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    my_list = list()

    # for i in range(sheet.ncols):
    #     for j in range(sheet.nrows):
    #         my_list.append(sheet.cell_value(j, i))

    for i in range(sheet.nrows):
        my_list.append(sheet.cell_value(i, 0))

    return my_list


def exportExcel(input_list):
    name_wanted = input("Name your excel file: ")
    excel_extension = '.xlsx'

    workbook = xlsxwriter.Workbook(name_wanted + excel_extension)
    worksheet = workbook.add_worksheet(dateTime())

    bold = workbook.add_format({'bold': True})

    worksheet.set_column('A:A', 47.86)
    worksheet.set_column('B:B', 10.43)
    worksheet.set_column('C:C', 9.86)
    worksheet.set_column('D:D', 8)
    # worksheet.set_column('F:F', 35.14)
    worksheet.set_column('F:F', 39.14)

    worksheet.write('A1', 'Name of Company', bold)
    worksheet.write('B1', 'Market', bold)
    worksheet.write('C1', 'Ticker', bold)
    worksheet.write('D1', 'Price', bold)
    worksheet.write('E1', 'Currency', bold)
    # worksheet.write('F1', 'Industry', bold)
    worksheet.write('F1', 'Time', bold)

    row = 1
    col = 0

    # for name, market, list_ticker, price, currency, industry, list_time in input_list:
    for name, market, list_ticker, price, currency, list_time in input_list:
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, market)
        worksheet.write(row, col + 2, list_ticker)
        worksheet.write(row, col + 3, price)
        worksheet.write(row, col + 4, currency)
        # worksheet.write(row, col + 5, industry)
        worksheet.write(row, col + 5, list_time)

        row += 1

    workbook.close()


def dateTime():
    current_date = str(date.today())
    # current_time = datetime.now().strftime("%H:%M:%S")
    return current_date


class Stock:

    def __init__(self, given_ticker):
        self.ticker = given_ticker
        url_starter = "http://finance.yahoo.com/quote/"
        self.url = url_starter + self.ticker
        result = requests.get(self.url)
        self.soup = BeautifulSoup(result.text, 'html.parser')

    def getPrice(self):
        price_line = self.soup.find("fin-streamer", {"class" : "Fw(b) Fz(36px) Mb(-4px) D(ib)"}).text
        price = float(price_line.replace(',', ''))
        return price

    def getCurrency(self):
        currency_line = self.soup.find("div", class_="C($tertiaryColor) Fz(12px)").text
        currency = currency_line[currency_line.find("Currency in ") + 11:].split()[0]
        return currency

    def getMarket(self):
        market_line = self.soup.find("div", class_="C($tertiaryColor) Fz(12px)").text
        market = market_line.split(' - ')[0]
        return market

    def getName(self):
        name_line = self.soup.find("h1", class_="D(ib) Fz(18px)").text
        name = name_line[0: name_line.find(" (")]
        return name

    # def getIndustry(self):
    #     industryLink = self.url + "/profile"
    #     profile_result = requests.get(industryLink)
    #     profile_soup = BeautifulSoup(profile_result.text, 'html.parser')
    #
    #     if profile_soup.find("span", attrs={"data-reactid": "23"}).text == "Industry":
    #         industry = profile_soup.find("span", class_="Fw(600)", attrs={"data-reactid": "25"}).text
    #     else:
    #         industry = profile_soup.find("span", class_="Fw(600)", attrs={"data-reactid": "27"}).text
    #
    #     return industry

    def ownList(self):
        listing = list()
        # listing.extend((self.getName(), self.getMarket(), self.ticker, self.getPrice(), self.getCurrency(),
        #                 self.getIndustry(), dateTime()))
        listing.extend((self.getName(), self.getMarket(), self.ticker, self.getPrice(), self.getCurrency(),
                        dateTime()))
        return listing

    def main(self):
        # return self.getName() + "'s Price: " + str(self.getPrice()) + " " + self.getCurrency() + ". Ticker: " + self.ticker \s
        #        + ". " + dateTime()

        return self.getName() + "'s Price: " + str(self.getPrice()) + " " + self.getCurrency() + ". Ticker: " + self.ticker \
               + ". " + dateTime()


print("Would you like to search up individual stocks or view spreadsheet? ")
print("[s] for individual stocks")
print("[v] to view spreadsheet")
print("[q] to quit")

while True:
    decision = input("Your choice: ")

    if decision == "s":
        while True:
            ticker = input("Enter the ticker/tickers (separated by space), or type q to quit: ")

            if ticker == "q":
                break
            else:
                temporary_list = ticker.split(" ")
                for ticker in temporary_list:
                    stock_initiate = Stock(ticker)
                    print(stock_initiate.main())
                    fullList.append(stock_initiate.ownList())
                    sleep(6) #delay for web scraping

                answer = input("Do you want to export these results? [y] for yes. ")
                if answer == "y":
                    exportExcel(fullList)
                else:
                    break

    elif decision == "v":

        list_created = usingExcel()

        for ticker in list_created:
            stock_initiate = Stock(ticker)
            print(stock_initiate.main())
            fullList.append(stock_initiate.ownList())
            sleep(6) #delay for web scraping

        answer = input("Do you want to export these results? [y] for yes. ")
        if answer == "y":
            exportExcel(fullList)
        else:
            break

        continue
    elif decision == "q":
        break
    else:
        print("Sorry, you have not entered an available option. Please enter again. ")
        continue
