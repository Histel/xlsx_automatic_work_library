from openpyxl import load_workbook
from openpyxl import Workbook

from mypackage.functions import search_name_for_value
from mypackage.functions import search_name_for_value_other
from mypackage.functions import set_column_headers
from mypackage.functions import add_change_values_cells
from mypackage.functions import return_values_cell

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

    all_values_list = return_values_cell(sheet, 'A')
    print(len(all_values_list))
    all_values_str = ''.join(all_values_list)
    new_all_values_list = re.split(';', all_values_str)

    users_list = []
    prevsymbol_slice = 0
    after_slice = 6
    for item in new_all_values_list:
        if not item:
            break
        users_list.append(new_all_values_list[prevsymbol_slice:after_slice])
        prevsymbol_slice += 7
        after_slice += 7

    admittance_dict = {}
    admittance = 0
    for users, next_user in zip_longest(users_list, users_list[1:], fillvalue=''):
            try:
                users_date = users[0][:2]
                next_users_date = next_user[0][:2]
                if users[5] == 'Допуск':
                    print('Допуск')
                    admittance += 1
                    admittance_dict[users[0][:2]] = admittance
                    if users_date < next_users_date:
                        admittance = 0
            except IndexError:
                pass

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