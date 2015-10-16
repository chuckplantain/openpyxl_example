from openpyxl import load_workbook
import datetime

#
## Testing variables
template = '/Users/kyleblazepetan/Plumb/masterxlsx/INTERNATIONAL PASSENGER MANIFEST_MASTER.xlsx'

trip_dict = {
u'date of trip': '12/24/2015',
u'trip number': 321456774,
u'generated on': '03/14/1998'
}

international = True

##
#

def write_xlsx_file(dict_list, trip_dict, template, international):
    '''
    :param dict_list: a list of dicts
    :param trip_dict: three item dict.
    :param template: uri for template xlsx file
    :param international: bool reflecting header style
    :return: True to indicate success
    '''
    wb = load_workbook(template)
    ws = wb.active
    header = []

    ws['C3'] = trip_dict[u'date of trip']
    ws['C4'] = trip_dict[u'trip number']
    ws['C8'] = trip_dict[u'generated on']

    if international:
        headerCells = ws['B11':'L11']
        cells = ws['B12':'L512']
    else:
        headerCells = ws['B11':'H11']
        cells = ws['B12':'H512']

    for row in headerCells:
        for cell in row:
            header.append(cell.value)

    for item, row in zip(dict_list, cells):
        for idx, cell in enumerate(row):
            if item[header[idx].lower()] is not None:
                cell.value = item[header[idx].lower()]
            else:
                continue

    wb.save('/Users/kyleblazepetan/Desktop/test.xlsx')
    return True
