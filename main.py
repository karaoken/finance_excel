from openpyxl import Workbook, load_workbook, worksheet

wb = load_workbook(filename="pl_conso.xlsx")

ws = wb["Sheet1"]

print(ws['A1'].value)


def get_statement(sheet: worksheet):
    stmt = sheet['B4'].value
    if stmt == "Income Statement":
        return "pl"
    elif stmt == "Balance Sheet":
        return "bs"
    elif stmt == "Cash Flow":
        return 'cf'


def get_symbol(sheet: worksheet):
    return sheet['B3'].value


def get_type(sheet: worksheet):
    return sheet['B5'].value


def delete_last_4_rows(sheet: worksheet):
    last_row = get_last_row(sheet, 1)
    sheet.delete_rows(last_row - 3, amount=4)


def get_last_row(sheet: worksheet, col):
    starting_row = 500
    for row in range(starting_row, 1, -1):
        if sheet.cell(row=row, column=col).value is not None:
            return row


def get_last_column(sheet: worksheet, row):
    starting_col = 500
    for col in range(starting_col, 1, -1):
        if sheet.cell(row=row, column=col).value is not None:
            return col


def convert_date_format(date):
    year = date[6:10]
    quarter = date[1:3]
    return f"{quarter}/{year}"


symbol = get_symbol(ws)
statement_type = get_type(ws)
statement = get_statement(ws)

# Copy to new worksheet
copied_worksheet = wb.copy_worksheet(ws)
new_sheet_name = f"{symbol}_{statement}_{statement_type}"
copied_worksheet.title = new_sheet_name
wb.worksheets.append(copied_worksheet)

new_sheet = wb[new_sheet_name]

delete_last_4_rows(new_sheet)
new_sheet.delete_rows(1, amount=12)
print(new_sheet['B1'].value)
print(convert_date_format(new_sheet['B1'].value))
print(get_last_column(new_sheet, 1))

# Change date format
last_col = get_last_column(new_sheet, 1)
last_row = get_last_row(new_sheet, 1)
for row in range(1, last_row + 1):
    print(new_sheet.cell(row=row, column=1).value)
for c in range(2, last_col + 1):
    converted_date = convert_date_format(new_sheet.cell(row=1, column=c).value)
    new_sheet.cell(row=1, column=c).value = converted_date
    print(convert_date_format(new_sheet.cell(row=1, column=c).value))
wb.save("pl_conso.xlsx")
