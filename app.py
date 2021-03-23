from openpyxl import load_workbook
from openpyxl import Workbook

from mypackage.functions import search_name_for_value
from mypackage.functions import search_name_for_value_other
from mypackage.functions import set_column_headers
from mypackage.functions import add_change_values_cells
from mypackage.functions import return_values_cell

from mypackage.app_functions import return_date_value_dict
from mypackage.app_functions import return_user_list
from mypackage.app_functions import split_all_values_list

# from tkinter import Tk
# from tkinter import Button
# from tkinter import filedialog
# from tkinter import Text
# from tkinter import Label

import threading
import time
import os
import csv
import sys
import re
from itertools import zip_longest


def main():
    book_data = load_workbook('zhurnal_medosmotrov_1_12-31_12.xlsx')
    sheet = book_data['Sheet']

    all_values_list = split_all_values_list(sheet, 'A')

    users_list = return_user_list(all_values_list)

    admittance_dict = return_date_value_dict(users_list, 'Допуск')

    print(admittance_dict)


if __name__ == '__main__':
    '''workbook = Workbook()
    worksheet = workbook.active
    with open('zhurnal_medosmotrov_1_12-31_12.csv', 'r') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                for idx, val in enumerate(col.split(',')):
                    cell = worksheet.cell(row=r + 1, column=c + 1)
                    cell.value = val
    workbook.save('zhurnal_medosmotrov_1_12-31_12.xlsx')'''
    main()