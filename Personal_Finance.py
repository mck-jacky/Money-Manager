from traceback import print_tb
from tracemalloc import start
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from currency_web_scrapping import get_currency
import datetime
from duplicate_transcation import duplicate_to_main
from stock import stock

INCOME = 0
EXPENDITURE = 1
AUD_COL = 9
HKD_COL = 10
CURRENCY_AMOUNT_COL = 9
ACCOUNTS_COL = 7
DATE_COL = 9

##################################### HELPER FUNCTIONS #####################################

def fill_in_income_and_expenditure(start_row, start_col, type):
    end_col_of_income_and_expenditure = 5
    count = 1

    if type == INCOME:
        number_of_type = int(input("Number of INCOME: "))
    elif type == EXPENDITURE:
        number_of_type = int(input("Number of EXPENDITURE: "))
    for row in range(start_row, start_row + number_of_type):
        for col in range(start_col, end_col_of_income_and_expenditure, 2):
            char = get_column_letter(col)
            if col == 2:
                name_of_type = input(f"Type {count}: ")
                count += 1
                ws[char + str(row)].value = name_of_type
            elif col == 4:
                amount = float(input(f"Amount of {name_of_type}? "))
                ws[char + str(row)].value = amount

def get_currency_col_char_and_set_amount(currency, row, amount):
    char = get_column_letter(currency)
    ws[char + str(row)].value = amount

def set_amount(start_row, end_row):
    char = get_column_letter(ACCOUNTS_COL)
    for row in range(start_row, end_row):
        account_name = ws[char + str(row)].value
        amount = float(input(f"$ in {account_name}: "))
        if "AUD" in account_name:
            currency_char = get_column_letter(AUD_COL)
            ws[currency_char + str(row)].value = amount
        elif "HKD" in account_name:
            currency_char = get_column_letter(HKD_COL)
            ws[currency_char + str(row)].value = amount
    
############################################################################################

wb = load_workbook("2022 Personal Finance.xlsx")

month_num = input("Which month are you going to modify? ")
datetime_object = datetime.datetime.strptime(month_num, "%m")
sheet_name = datetime_object.strftime("%B")
ws = wb[sheet_name]

################################### INCOME & EXPENDITURE ###################################

start_row_of_income = 2
start_col_of_income = 2
start_row_of_expenditure = 14
start_col_of_expenditure = 2

fill_in_income_and_expenditure(start_row_of_income, start_col_of_income, INCOME)
print("\n")
fill_in_income_and_expenditure(start_row_of_expenditure, start_col_of_expenditure, EXPENDITURE)

############################################################################################
print("\n")
########################################## ASSETS ##########################################
AUD_COL = 9
HKD_COL = 10

# CASH
start_row_of_cash = 3
end_row_of_cash = 5

for row in range(start_row_of_cash, end_row_of_cash):
    if row == start_row_of_cash:
        amount_cash = float(input("AUD in Cash: "))
        char = get_column_letter(AUD_COL)
        ws[char + str(row)].value = amount_cash
    elif row == start_row_of_cash + 1:
        amount_cash = float(input("HKD in Cash: "))
        char = get_column_letter(HKD_COL)
        ws[char + str(row)].value = amount_cash
    
# ACCOUNTS
start_row_of_accounts = 6
end_row_of_accounts = 10
set_amount(start_row_of_accounts, end_row_of_accounts)

# Saving
start_row_of_saving = 13
end_row_of_saving = 15
set_amount(start_row_of_saving, end_row_of_saving)

############################################################################################
print("\n")
###########################################STOCK############################################
STOCK_NAME = 13
STOCK_QTY = 14
STOCK_PRICE = 15

start_row_of_stock = 3

number_of_stocks = int(input("Number of stocks?: "))
for row in range(start_row_of_stock, start_row_of_stock + number_of_stocks):
    stock_code = input("Stock code: ")
    stock_qty = float(input('Qty: '))

    code_char = get_column_letter(STOCK_NAME)
    qty_char = get_column_letter(STOCK_QTY)
    price_char = get_column_letter(STOCK_PRICE)

    ws[code_char + str(row)].value = stock_code
    ws[qty_char + str(row)].value = stock_qty
    ws[price_char + str(row)].value = stock(stock_code)

############################################################################################
print("\n")
###################################CURRENCY(WEB SCRAPPING)##################################
start_row_of_currency = 23
end_row_of_currency = 25

for row in range(start_row_of_currency, end_row_of_currency):
    if row == start_row_of_currency:
        char = get_column_letter(CURRENCY_AMOUNT_COL)
        ws[char + str(row)].value = get_currency("https://themoneyconverter.com/AUD/HKD")
        char = get_column_letter(15)
        ws[char + str(row)].value = get_currency("https://themoneyconverter.com/USD/AUD")
    elif row == start_row_of_currency + 1:
        char = get_column_letter(CURRENCY_AMOUNT_COL)
        ws[char + str(row)].value = get_currency("https://themoneyconverter.com/HKD/AUD")



today = datetime.date.today()
date = today.strftime("%d/%m/%Y")
char = get_column_letter(DATE_COL)
ws[char + str(25)].value = date

############################################################################################

################################### Duplicate Transcation ##################################
transcation_wb = load_workbook("2022-01-01 _ 01-31.xlsx")
transcation_ws = transcation_wb.active
main_transcation_ws = wb["Transcation Details"]
duplicate_to_main(transcation_ws, main_transcation_ws)

############################################################################################

print("Saved")
wb.save("2022 Personal Finance.xlsx")


                