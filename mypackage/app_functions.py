from openpyxl import load_workbook
from openpyxl import Workbook

from mypackage.functions import search_name_for_value
from mypackage.functions import search_name_for_value_other
from mypackage.functions import set_column_headers
from mypackage.functions import add_change_values_cells
from mypackage.functions import return_values_cell

from itertools import zip_longest
import re


def split_all_values_list(sheet: 'ws.active', cell_name: 'example "A"') -> list:
    ''' Возвращает список из разбитых элементов через ";" '''
    all_values_list = return_values_cell(sheet, cell_name)
    print(len(all_values_list))
    all_values_str = ''.join(all_values_list)
    new_all_values_list = re.split(';', all_values_str)
    return new_all_values_list


def return_user_list(all_values_list: 'list with data') -> list:
    ''' Возвращает список со всеми значениями разбив их на отдельные списки (по 6 элементов в каждом) '''
    user_list = []
    prevsymbol_slice = 0
    after_slice = 6
    for item in all_values_list:
        if not item:
            break
        user_list.append(all_values_list[prevsymbol_slice:after_slice])
        prevsymbol_slice += 7
        after_slice += 7
    return user_list


def return_date_value_dict(
        data_list: 'data',
        value: 'for search',
        ) -> dict:
    ''' Возвращает словарь даты (в формате 01, 02 *день месяца*) + количество значения '''
    admittance_dict = {}
    admittance = 0
    for users, next_user in zip_longest(data_list, data_list[1:], fillvalue=''):
        try:
            users_date = users[0][:2]
            next_users_date = next_user[0][:2]
            if users[5] == value:
                admittance += 1
                admittance_dict[users[0][:2]] = admittance
                if users_date < next_users_date:
                    admittance = 0
        except IndexError:
            pass
    return admittance_dict