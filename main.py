import openpyxl as xl


def save_file(workbook, file_name):
    workbook.save(file_name)


def calculate_all_income(sheet):
    for i in range(2, sheet.max_row + 1):
        hour = sheet.cell(i, 4).value
        income_hour = sheet.cell(i, 5).value
        all_income = hour * income_hour
        print(f'all_income is {all_income}')
        sheet.cell(i, 6).value = all_income


def load_file(file_name):
    wb = xl.load_workbook(file_name)
    print(f'type {wb}')
    sheet = wb['Sheet1']
    calculate_all_income(sheet)
    save_file(workbook=wb, file_name=file_name)


file_name = 'info_employee.xlsx'
load_file(file_name=file_name)
