from openpyxl import load_workbook
from mypackage.functions import search_name_for_value, search_name_for_value_other


def main():

    excel_document = load_workbook('test.xlsx')

    sheet = excel_document['Sheet1']

    turnout = search_name_for_value(sheet, 'B', 'I', 'Москва')
    print(turnout)

    turnout_and_value = search_name_for_value_other(sheet, 'C', 'G', 'I', 'Москва')
    print(turnout_and_value)


if __name__ == '__main__':
    main()