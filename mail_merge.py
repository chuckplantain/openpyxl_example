from openpyxl import load_workbook
import re

def process_merge(data, flightSegmentData):
    """
    :param data: a dict of values, keys being the <<Values>> in lowercase form
    :param flightSegmentData: array of dicts corresponding to flight legs
    :return: True if all went well
    """
    pattern = re.compile(r'(<<)(\w+)(>>)')
    wb = load_workbook('/Users/kyleblazepetan/Desktop/mailMergeTemplate.xlsx')
    ws = wb.active

    def zipForm(segments):
        pertinent_rows = ws['A32':'L54']
        for templateRow, blankRow in zip(pertinent_rows, ws.iter_rows(row_offset=start_row)):
            for template_cell, current_cell in zip(templateRow, blankRow):
                copied_style = template_cell.style
                current_cell.value = template_cell.value
                current_cell.style = copied_style


    def buildForm(segments):
        pertinent_rows = ws['A32':'L54']
        start = 56
        for i in range(1, segments + 2):
            for rowidx, rows in enumerate(pertinent_rows, start=start):
                for colidx, template_cell in enumerate(rows, start=1):
                    current_cell = ws.cell(row = rowidx, column = colidx)
                    copied_style = template_cell.style
                    current_cell.value = template_cell.value
                    current_cell.style = copied_style
                    start += 25

    start_row = 33
    flightSegments = len(flightSegmentData)
    while flightSegments > 1:
        start_row += 25
        zipForm(flightSegments)
        flightSegments -= 1

#    def lookUpFunction(matchobj):
#        resp = data[str(matchobj.group(2)).lower()]
#        if resp:
#            return resp
#        return 'key not found'
#
#    for row in ws.iter_rows():
#        for cell in row:
#            if cell.value is not None:
#                cell.value = re.sub(pattern, lookUpFunction, cell.value.lower())
#
    wb.save('/Users/kyleblazepetan/Desktop/mergeTesting.xlsx')
    return True
