import re

PATTERN = '0|1|2|3|4|5|6|7|8|9'


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

        cell = re.sub(PATTERN, '', cell)
        cell_names = re.sub(PATTERN, '', cell_names)
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
    names_and_phones_dict = {}
    for i in range(1, sheet.max_row + 1):
        cell_values = cell_values + str(i)
        cell_other_values = cell_other_values + str(i)
        cell_names = cell_names + str(i)
        prevsymbol_other_values = sheet[cell_other_values].value
        prevsymbol_names = sheet[cell_names].value
        prevsymbol_values = sheet[cell_values].value
        if prevsymbol_values == value:
            names_and_phones_dict[prevsymbol_names] = prevsymbol_other_values

        cell_values = re.sub(PATTERN, '', cell_values)
        cell_other_values = re.sub(PATTERN, '', cell_other_values)
        cell_names = re.sub(PATTERN, '', cell_names)
    return names_and_phones_dict


def _test():
    assert as_int('1') == 1

if __name__ == "__main__":
    _test()
