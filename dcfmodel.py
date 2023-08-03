import os
from sys import argv
import requests
from openpyxl import load_workbook, Workbook
import shutil
from sec_api import QueryApi, XbrlApi
import yfinance as yf
import pandas as pd

#API key for SEC API
api_key="insert_your_api_key_here"

# Read command line arguments
def main():
    if len(argv) < 4:
        print("Insufficient arguments provided.")
        print("Usage: python sec_api.py [ticker] [start_year] [end_year]")
    elif argv[2] == argv[3]:
        print("Program requires more than one year's worth of data.")
        sys.exit()
    else:
        ticker = argv[1]
        start_year = argv[2]
        end_year = argv[3]
    
    create_excel_copy(ticker)
    fetchandinsert(ticker, start_year, end_year)


def create_excel_copy(ticker):
    # Check if the template file exists
    if not os.path.isfile("template.xlsx"):
        print("Template file not found.")
        return

    # New file name for the copy
    new_file_name = f"{ticker}.xlsx"
    template_file = "template.xlsx"
    
    # Copy the template file to the new file
    shutil.copyfile(template_file, new_file_name)
    
    print("Excel file created:", new_file_name)


def fetchandinsert(ticker, start_year, end_year):
    queryApi = QueryApi(api_key)

    # Use SEC query_api to get all filings of a company filed between the periods.
    query = {
        "query": { "query_string": {
        "query": "ticker:{} AND filedAt:{{ {}-01-01 TO {}-12-31 }} AND formType:\"10-K\"".format(ticker, start_year, end_year)
        } },
    "from": "0",
    "size": "10",
    "sort": [{ "filedAt": { "order": "desc" } }]
    }

    data = queryApi.get_filings(query)
    filings = data['filings']

    num_filings = sum(1 for filing in filings if filing['formType'] == '10-K')
    print(f"Number of filings: {num_filings}")
    
    # Insert columns (already space for one filing in template so no need to insert all)
    num_columns = num_filings - 1

    # Start with earlier years
    filings.reverse()

    # Load new file using openpyxl 
    filename = f"{ticker}.xlsx"
    workbook = load_workbook(filename)
    sheet = workbook["Model"]
    main = workbook["Main"]
    
    # Insert columns starting at column C (3rd column)
    sheet.insert_cols(3, num_columns)

    # Retrieve and insert price and beta from yahoo finance
    stock = yf.Ticker(ticker)
    current_price = stock.info.get("regularMarketPreviousClose")
    market_cap = stock.info.get('marketCap')
    beta = stock.info.get('beta')
    main['C3'] = current_price
    main['F4'] = beta
    
    # Get most recent years interest expense for cost of debt calculation - yfis stands for yahoo finance income statement
    yfis = stock.income_stmt
    interest_expense = yfis.iloc[yfis.index == 'Interest Expense', 0].values[0]

    # Iterate over filings and write information into the inserted columns
    i = 0
    for filing in filings:
        if filing['formType'] != '10-K':
            continue
    
        url = filing['linkToFilingDetails']
        xbrlApi = XbrlApi(api_key)
        xbrl_json = xbrlApi.xbrl_to_json(htm_url=url)

        # Check if the current filing is the last '10-K' filing - use most recent information for this part (Shares, Cash, Debt)
        if filing == filings[-1]:
            Shares = int(xbrl_json['StatementsOfIncome'].get('WeightedAverageNumberOfSharesOutstandingBasic', [{'value': '0'}])[0]['value'])
            Cash = int(xbrl_json['BalanceSheets'].get('CashAndCashEquivalentsAtCarryingValue', [{'value': '0'}])[0]['value']) +\
                int(xbrl_json['BalanceSheets'].get('ShortTermInvestments', [{'value': '0'}])[0]['value'])
            Debt = int(xbrl_json['BalanceSheets'].get('LongTermDebtCurrent', [{'value': '0'}])[0]['value']) +\
                int(xbrl_json['BalanceSheets'].get('LongTermDebtNoncurrent', [{'value': '0'}])[0]['value'])

            main['C6'] = Cash
            main['C7'] = Debt
            main['C4'] = Shares
            main['I6'] = interest_expense
        
        # Only use the information from the highest year to avoid overlap
        years = [x['period']['endDate'].split('-')[0] for x in xbrl_json['StatementsOfIncome']['NetIncomeLoss']]

        highest_year = max(years)
        executed = False

        
        for item in xbrl_json['StatementsOfIncome']['NetIncomeLoss']:
            date = item['period']['endDate']
            year = int(date.split('-')[0])

            if year == int(highest_year) and not executed:
                executed = True

                # Store the information in variables
                date = item['period']['endDate']
                year = date.split("-")[0]

                # Retrieve the following values associated with the highest year - If value is under different heading return 0
                Revenue = int(xbrl_json['StatementsOfIncome'].get('RevenueFromContractWithCustomerExcludingAssessedTax', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('Revenues', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('SalesRevenueNet', [{'value': '0'}])[0]['value'])
                COGS = int(xbrl_json['StatementsOfIncome'].get('CostOfGoodsAndServicesSold', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('CostRevenue', [{'value': '0'}])[0]['value'])
                RandD = int(xbrl_json['StatementsOfIncome'].get('ResearchAndDevelopmentExpense', [{'value': '0'}])[0]['value'])
                GA = int(xbrl_json['StatementsOfIncome'].get('SellingGeneralAndAdministrativeExpense', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('GeneralAndAdministrativeExpense', [{'value': '0'}])[0]['value'])
                SandM = int(xbrl_json['StatementsOfIncome'].get('SellingAndMarketingExpense', [{'value': '0'}])[0]['value'])
                Restructuring = int(xbrl_json['StatementsOfIncome'].get('RestructuringAndOtherExpenses', [{'value': '0'}])[0]['value']) + int(xbrl_json['StatementsOfIncome'].get('RestructuringCharges', [{'value': '0'}])[0]['value'])
                OpEx = int(xbrl_json['StatementsOfIncome'].get('OperatingExpenses', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('CostsAndExpenses', [{'value': '0'}])[0]['value'])
                OPinc = int(xbrl_json['StatementsOfIncome'].get('OperatingIncomeLoss', [{'value': '0'}])[0]['value'])
                interest = int(xbrl_json['StatementsOfIncome'].get('InterestIncomeExpenseNonoperatingNet', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('InvestmentIncomeInterest', [{'value': '0'}])[0]['value']) - \
                    int(xbrl_json['StatementsOfIncome'].get('InterestExpense', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('InvestmentIncomeNonoperating', [{'value': '0'}])[0]['value'])
                othernonop = int(xbrl_json['StatementsOfIncome'].get('NonoperatingIncomeExpense', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest', [{'value': '0'}])[0]['value']) + \
                    int(xbrl_json['StatementsOfIncome'].get('OtherNonoperatingIncomeExpense', [{'value': '0'}])[0]['value'])
                Taxes = int(xbrl_json['StatementsOfIncome'].get('IncomeTaxExpenseBenefit', [{'value': '0'}])[0]['value'])
                
                net_income = int(xbrl_json['StatementsOfIncome'].get('NetIncomeLoss', [{'value': '0'}])[0]['value'])
                Shares = int(xbrl_json['StatementsOfIncome'].get('WeightedAverageNumberOfSharesOutstandingBasic', [{'value': '0'}])[0]['value'])
                EPS = float(xbrl_json['StatementsOfIncome'].get('EarningsPerShareBasic', [{'value': '0'}])[0]['value'])
                EPS_diluted = float(xbrl_json['StatementsOfIncome'].get('EarningsPerShareDiluted', [{'value': '0'}])[0]['value'])


                column_letter = str(chr(67 + i))  # Starts from 'C' (66 in ASCII)

                Gross_Profit = Revenue - COGS
                
                # Write information into the corresponding column (divided by 1000)
                sheet[column_letter + '1'] = year
                
                sheet[column_letter + '6'] = Revenue
                sheet[column_letter + '8'] = COGS
                sheet[column_letter + '10'] = Gross_Profit
                sheet[column_letter + '11'] = RandD
                sheet[column_letter + '12'] = SandM
                sheet[column_letter + '13'] = GA
                sheet[column_letter + '14'] = Restructuring
                sheet[column_letter + '15'] = OpEx if OpEx != 0 else (RandD + GA + SandM + Restructuring) #incase opex is not on a companys income statement
                sheet[column_letter + '16'] = Gross_Profit - OpEx
                sheet[column_letter + '17'] = interest
                sheet[column_letter + '18'] = othernonop
                sheet[column_letter + '20'] = Taxes
                sheet[column_letter + '21'] = net_income
                sheet[column_letter + '24'] = Shares
                sheet[column_letter + '25'] = EPS
                sheet[column_letter + '26'] = EPS_diluted

                print(f"Inserted data for {year}")
                i+=1
    
    # Format columns
    actcol= str(chr(68 + num_columns)) # (D) by default - actual column to be edited
    refcol = str(chr(67 + num_columns)) # (C) by default - column being referencened in formula
    for i in range(7): # Seven years- from 2023 to 2029
        sheet[actcol + '6'] = '=' + refcol + '6*(1+$B$31)' #revenue
        sheet[actcol + '8'] = '=' + actcol + '6-' + actcol + '10' #cogs
        sheet[actcol + '9'] = '=' + refcol + '9*(1+$B$32)' # gross margin
        sheet[actcol + '10'] = '=' + actcol + '6*' + actcol + '9'
        sheet[actcol + '11'] = '=' + refcol + '11*(1+$B$33)' #r&d
        sheet[actcol + '12'] = '=' + refcol + '12*(1+$B$34)' #s&m
        sheet[actcol + '13'] = '=' + refcol + '13*(1+$B$35)' #g&a
        sheet[actcol + '14'] = '=' + refcol + '14*(1+$B$36)' #restructuring
        sheet[actcol + '15'] = '=' + actcol + '11+' + actcol + '12+' + actcol + '13+' + actcol + '14' #opex
        sheet[actcol + '16'] = '=' + actcol + '10-' + actcol + '15' #opinc
        sheet[actcol + '17'] = '=' + refcol + '17*(1+$B$37)' #interest
        sheet[actcol + '18'] = '=' + refcol + '18*(1+$B$38)' #other nonop
        sheet[actcol + '19'] = '=' + actcol + '16+' + actcol + '17+' + actcol + '18' #pretax income
        sheet[actcol + '20'] = '=' + refcol + '20*(1+$B$39)' #taxes
        sheet[actcol + '21'] = '=' + actcol + '19-' + actcol + '20' #net income
        sheet[actcol + '24'] = '=' + refcol + '24*(1+$B$40)' #shares

        sheet[actcol + '22'] = '=' + actcol + '21/(1+Main!$I$9)^' + str(i+1) #npv
        
        actcol = str(chr(ord(actcol) + 1))
        refcol = str(chr(ord(refcol) + 1))
    
    # Revenue year on year formula
    d_column = str(chr(68))
    c_column = str(chr(67))
    formatall = 6 + num_columns
    for j in range(formatall):
        sheet[d_column + '7'] = '=' + d_column + '6/' + c_column + '6-1' 
        d_column = str(chr(ord(d_column)+1))
        c_column = str(chr(ord(c_column)+1))

    # Gross margin and pretax income formula
    c_column = str(chr(67))
    for k in range(num_columns + 1):
        sheet[c_column + '9'] = '=' + c_column + '10/' + c_column + '6' #gross margin
        sheet[c_column + '19'] = '=' + c_column + '16+' + c_column + '17+' + c_column + '18' #pretax income
        c_column = str(chr(ord(c_column)+1))

    # Final forecast year in L3 (terminal value calculation) and npv value in C15
    final_col= str(chr(74 + num_columns))
    tv_col = str(chr(75 + num_columns))
    first_col = str(chr(68 + num_columns))
    main['L3'] = '=Model!' + final_col + '21'
    main['C15'] = '=Model!' + tv_col + '23'
    sheet[tv_col + '23'] = '=SUM(' + first_col + '22:' + tv_col + '22)'
    sheet[tv_col + '22'] = '=' + tv_col + '21/(1+Main!$I$9)^7'

    print("File formatted")
    workbook.save(filename)



if __name__ == '__main__':
    main()

