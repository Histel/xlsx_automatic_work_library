import re

NUMS_CLEAR = '0|1|2|3|4|5|6|7|8|9'


def return_values_cell(
        sheet: 'ws.active',
        cell_name: 'your names cell (example: "B")',
        ) -> list:

    return_list = []
    for i in range(2, sheet.max_row + 1):
        cell_name = cell_name + str(i)
        prevsymbol = sheet[cell_name].value
        return_list.append(prevsymbol + ';')
        cell_name = re.sub(NUMS_CLEAR, '', cell_name)
    return return_list

def add_change_values_cells(
        sheet: 'ws.active',
        **cell_name: 'names A4, B54, C32.., example: (B21="text", or B="text")'
        ) -> 'add/change value in xlsx':

    ''' Adds / Modifies a value in the specified table '''
    for cell, text in cell_name.items():
        cell = str(cell)
        if len(cell) == 1:
            cell = cell + '1'
            sheet[cell] = text
        else:
            sheet[cell] = text
    return None


def set_column_headers(
        sheet: 'ws.active',
        **kwargs: 'names A1, B1, C1.., example: (B="text")'
        ) -> 'print in xlsx':

    ''' sets headings to the first cells of columns '''
    for cell, text in kwargs.items():
        cell = str(cell)
        if len(cell) >= 2:
            raise ValueError('the value must contain only one of the letters in the column (A, B, C, D..).')
        else:
            cell = cell + '1'
            sheet[cell] = text
    return None


def search_name_for_value(
        sheet: 'ws.active',
        cell_names: 'your names cell (example: "B")',
        cell: 'values cell search',
        value: 'your search value'
        ) -> list:

    ''' looks for a specific value in a table and
    returns a list of names * complete '''
    names_list = []
    for i in range(1, sheet.max_row+1):
        cell = cell + str(i)
        cell_names = cell_names + str(i)
        prevsymbol_names = sheet[cell_names].value
        prevsymbol = sheet[cell].value
        if prevsymbol == value:
            names_list.append(prevsymbol_names)

        cell = re.sub(NUMS_CLEAR, '', cell)
        cell_names = re.sub(NUMS_CLEAR, '', cell_names)
    names_set = set(names_list)
    return names_set


def search_name_for_value_other(
        sheet: 'ws.active',
        cell_names: 'your names cell (example: "B")',
        cell_other_values: 'other values cell (example: "phone")',
        cell_values: 'values cell search',
        value: 'your search value'
        ) -> dict:

    ''' looks for a specific value in one table, and returns 2 others
    needed in the dictionary (value: value) * complete '''
    names_and_values_dict = {}
    for i in range(1, sheet.max_row + 1):
        cell_values = cell_values + str(i)
        cell_other_values = cell_other_values + str(i)
        cell_names = cell_names + str(i)
        prevsymbol_other_values = sheet[cell_other_values].value
        prevsymbol_names = sheet[cell_names].value
        prevsymbol_values = sheet[cell_values].value
        if prevsymbol_values == value:
            names_and_values_dict[prevsymbol_names] = prevsymbol_other_values

        cell_values = re.sub(NUMS_CLEAR, '', cell_values)
        cell_other_values = re.sub(NUMS_CLEAR, '', cell_other_values)
        cell_names = re.sub(NUMS_CLEAR, '', cell_names)
    return names_and_values_dict


def _test():
    assert as_int('1') == 1

if __name__ == "__main__":
    _test()
