from openpyxl import load_workbook
import re

def mail_merge(leg_dict, trip_dict):
    wb = load_workbook('/Users/kyleblazepetan/Desktop/mailMergeTemplate.xlsx')
    ws = wb.active
    handlebar_pattern = re.compile(r'(<<)(\w+)(>>)')
    flight_segment_pattern = re.compile(r'FLIGHT SEGMENT ')
    segments = (len(leg_dict))
    max_value = 55 + ( 23 * segments ) - 23
    for index in xrange(55, max_value, 23):
        segment_count = (index - 55) / 23 + 2
        def legLookUpFunction(matchobj):
            resp = leg_dict[segment_count - 1][str(matchobj.group(2)).lower()]
            if resp:
                return resp
            return 'key not found'
        blank_rows = ws.iter_rows(row_offset=index)
        pertinent_rows = ws['A32':'L54']
        for pertinent_row, blank_row in zip(pertinent_rows, blank_rows):
            for pertinent_cell, blank_cell in zip(pertinent_row, blank_row):
                desired_style = pertinent_cell.style
                blank_cell.style = desired_style
                if pertinent_cell.value is not None:
                    val = pertinent_cell.value
                    if flight_segment_pattern.search(val) is not None:
                        blank_cell.value = "Flight Segment " +  str(segment_count)
                    elif handlebar_pattern.search(val) is not None:
                        blank_cell.value = re.sub(handlebar_pattern, legLookUpFunction, pertinent_cell.value.lower())
                    else:
                        blank_cell.value = pertinent_cell.value
                else:
                    blank_cell.value = pertinent_cell.value

    def tripLookUpFunction(matchobj):
        resp = trip_dict[str(matchobj.group(2)).lower()]
        if resp:
            return resp
        return 'key not found'

    for row in ws.iter_rows('A1:L31'):
        for cell in row:
            if cell.value is not None:
                cell.value = re.sub(handlebar_pattern, tripLookUpFunction, cell.value.lower())

    wb.save('foo.xlsx')
