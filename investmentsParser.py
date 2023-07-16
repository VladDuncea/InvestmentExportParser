import sys
import csv
from openpyxl import Workbook, load_workbook, utils
import datetime


def extractDateFromDateTime(date: datetime.datetime)-> datetime.datetime:
    return datetime.datetime(year=date.year, month=date.month, day=date.day)

def compareDates(date1: datetime.datetime, date2: datetime.datetime):
    return date1.strftime("%Y-%m-%d") == date2.strftime("%Y-%m-%d")

class InvestmentParser:
    def __init__(self, filePath: str, type: str):
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
        self.resultXls.create_sheet("Taxes+Comissions")


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
            self.handleXtbCashHistRow(row)

        # Define variable to read sheet
        sheetClosedOp = excelFile["CLOSED POSITION HISTORY"]

        # Iterate over the rows and col
        for row in sheetClosedOp.iter_rows(14, sheetClosedOp.max_row):
            self.handleXtbClosedOpRow(row)

        self.exportResult("xtb")

    def handleXtbClosedOpRow(self, row):
        try:
            transactDateOpen = extractDateFromDateTime(row[5].value)
            transactDateClose = extractDateFromDateTime(row[7].value)
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


    def handleXtbCashHistRow(self, row):
        try:
            transactType = row[2].value
            transactDate = extractDateFromDateTime(row[3].value)
            transactComment = row[4].value
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
        elif transactType in ['tax RO', 'SEC fee']:
            self.cacheDict['taxes_comissions'].append({"date":transactDate, "value":value, "type":transactType, "moreInfo":transactComment})
        elif transactType in ['Stocks/ETF purchase', 'Profit/Loss', 'Stocks/ETF sale']:
            # nothing to do yet
            pass
        else:
            print(f"WARN: Unknown transaction type: {transactType}")

    def exportResult(self, filePrefix: str):
        # Dividend sheet
        sheet = self.resultXls["Dividends"]
        sheet.append(["Date","Company", "Value"])
        for idx, row in enumerate(self.cacheDict['dividends'],2):
            sheet.append([row["date"], row["company"], row["value"]])
            sheet[f'A{idx}'].number_format = 'dd-mm-yy'
            sheet[f'C{idx}'].number_format = '_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* "-"??_);_(@_)'

        # Deposits sheet
        sheet = self.resultXls["Deposits"]
        sheet.append(["Date", "Value"])
        for idx, row in enumerate(self.cacheDict['deposits'],2):
            sheet.append([row["date"], row["value"]])
            sheet[f'A{idx}'].number_format = 'dd-mm-yy'
            sheet[f'B{idx}'].number_format = '_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* "-"??_);_(@_)'

        # Sales sheet
        sheet = self.resultXls["Sales"]
        sheet.append(["Company", "Open Date", "Sell Date", "Buy Value", "Sell Value", "Profit"])
        for idx, row in enumerate(self.cacheDict['sales'],2):
            sheet.append([row["company"], row["dateOpen"],row["dateClose"],
                          row["openValue"],row["closeValue"],
                          row["closeValue"]-row["openValue"]])
            sheet[f'B{idx}'].number_format = 'dd-mm-yy'
            sheet[f'C{idx}'].number_format = 'dd-mm-yy'
            sheet[f'D{idx}'].number_format = '_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* "-"??_);_(@_)'
            sheet[f'E{idx}'].number_format = '_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* "-"??_);_(@_)'
            sheet[f'F{idx}'].number_format = '_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* "-"??_);_(@_)'

        # Taxes and comissions sheet
        sheet = self.resultXls["Taxes+Comissions"]
        sheet.append(["Reason", "Date", "Value", "Comment"])
        for idx, row in enumerate(self.cacheDict['taxes_comissions'],2):
            sheet.append([row["type"], row["date"],row["value"], row["moreInfo"]])
            sheet[f'B{idx}'].number_format = 'dd-mm-yy'
            sheet[f'C{idx}'].number_format = '_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* "-"??_);_(@_)'

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