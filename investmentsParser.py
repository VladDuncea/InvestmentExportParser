import sys
import csv
from openpyxl import Workbook, load_workbook
import datetime

class InvestmentParser:
    def __init__(self, filePath, type):
        self.filePath = filePath
        self.type = type
        # Init result xls file
        self.initResultXls()
        self.cacheDict = {"dividends":[], "deposits": [], "sales": [], "taxes_comissions":[]}

    def initResultXls(self):
        self.resultXls = Workbook()
        sheet = self.resultXls.active
        # rename initial sheet
        sheet.title = "Dividends"
        # create sheet for deposits
        self.resultXls.create_sheet("Deposits")
        # create sheet for sales
        self.resultXls.create_sheet("Sales")
        # create sheet for taxes/comissions
        self.resultXls.create_sheet("Taxes&Comissions")


    def parse(self):
        if self.type == 'xtb':
            self.parseXtb()
        elif self.type == 'etoro':
            raise NotImplementedError
        elif self.type == 'revolut':
            raise NotImplementedError
        else:
            print(f"ERR: Wrong type of file, expected 'xtb', 'etoro', 'revolut', got '{self.type}'\n")


    def parseXtb(self):
        # Define variable to load the dataframe
        excelFile = load_workbook(self.filePath)

        # Define variable to read sheet
        sheetCashOp = excelFile["CASH OPERATION HISTORY"]

        # Iterate over the rows and col
        for row in sheetCashOp.iter_rows(12, sheetCashOp.max_row):
            self.handleXtbDivRow(row)

        # Define variable to read sheet
        sheetClosedOp = excelFile["CLOSED POSITION HISTORY"]

        # Iterate over the rows and col
        for row in sheetClosedOp.iter_rows(14, sheetClosedOp.max_row):
            self.handleXtbDivRow(row)

        self.exportResult("xtb")

    def handleXtbClosedOpRow(self, row):
        try:
            transactDateOpen = row[5].value.strftime('%Y-%m-%d')
            transactDateClose = row[7].value.strftime('%Y-%m-%d')
            transactSymbol = row[2].value
            openValue = row[11].value
            closeValue = row[12].value
        except:
            # maybe out of data range
            return

        self.cacheDict['sales'].append({"dateOpen":transactDateOpen,
                                        "dateClose": transactDateClose,
                                        "company":transactSymbol,
                                        "openValue":openValue,
                                        "closeValue":closeValue,
                                        })


    def handleXtbDivRow(self, row):
        try:
            transactType = row[2].value
            transactDate = row[3].value.strftime('%Y-%m-%d')
            transactSymbol = row[5].value # not all rows have a symbol
            value = row[6].value
        except:
            # maybe out of data range
            return

        if transactType in ['Dividend', 'Withholding tax', 'Stamp duty']:
            # if len(self.cacheDict['dividends']) > 0:
            #     print(len(self.cacheDict['dividends']) > 0, self.cacheDict['dividends'][-1]["date"] == transactDate, self.cacheDict['dividends'][-1]["company"] == transactSymbol)
            if len(self.cacheDict['dividends']) > 0 and self.cacheDict['dividends'][-1]["date"] == transactDate and self.cacheDict['dividends'][-1]["company"] == transactSymbol:
                self.cacheDict['dividends'][-1]["value"] += value
            else:
                self.cacheDict['dividends'].append({"date":transactDate, "company":transactSymbol, "value":value})
        elif transactType == 'Deposit':
            self.cacheDict['deposits'].append({"date":transactDate, "value":value})
        elif transactType in ['Stocks/ETF purchase', 'Profit/Loss', 'Stocks/ETF sale']:
            # nothing to do yet
            pass
        else:
            print(f"WARN: Unknown transaction type: {transactType}")

    def exportResult(self, filePrefix):
        # Dividend sheet
        sheet = self.resultXls["Dividends"]
        sheet.append(["Date","Company", "Value"])
        for row in self.cacheDict['dividends']:
            sheet.append([row["date"], row["company"], row["value"]])
        # Deposits sheet
        sheet = self.resultXls["Deposits"]
        sheet.append(["Date", "Value"])
        for row in self.cacheDict['deposits']:
            sheet.append([row["date"], row["value"]])
        # Sales sheet
        sheet = self.resultXls["Sales"]
        # export parse result
        self.resultXls.save(f"{filePrefix}_investments_{datetime.datetime.now().strftime('%Y_%m_%d')}.xlsx")



# takes arguments form command line, expects 2 args, the path of the excel file and the type of import (xtb or etoro or revolut)
def main(args):
    if len(args) != 3:
        print(f"ERR: Wrong number of parameters! Got {len(args)-1}, but expecting 2 parameters: <path-to-excel> <type-of-import>\n")

    # Params
    filePath = args[1]
    type = args[2]
    # create the class
    investmentParser = InvestmentParser(filePath, type)
    investmentParser.parse()


main(sys.argv)