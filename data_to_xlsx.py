from openpyxl import load_workbook
from manifest import process_manifest
import datetime

output = '/Users/kyleblazepetan/Plumb/xlsxFiles/INTERNATIONAL PASSENGER MANIFEST.xlsx'
data = [{u'date of birth': datetime.date(2013, 3, 22),
  u'expiration date': datetime.date(2015, 12, 31),
  u'first name': u'Kyle',
  u'issue date': datetime.date(2015, 10, 1),
  u'issuing country': u'United States of America',
  u'last name': u'Petan',
  u'middle initial': u'B',
  u'passport number': 192,
  u'sex': u'M',
  u'weight': 100190,
  u'weight_unit': u'kg'},
 {u'date of birth': datetime.date(1981, 1, 1),
  u'expiration date': datetime.date(2015, 12, 31),
  u'first name': u'Karen',
  u'issue date': datetime.date(2015, 10, 1),
  u'issuing country': u'United States of America',
  u'last name': u'Seleb',
  u'middle initial': u'A',
  u'passport number': 334,
  u'sex': u'F',
  u'weight': 200,
  u'weight_unit': u'lbs'},
 {u'date of birth': datetime.date(1999, 3, 5),
  u'expiration date': datetime.date(2015, 12, 31),
  u'first name': u'Joe',
  u'issue date': datetime.date(2001, 10, 5),
  u'issuing country': u'United States of America',
  u'last name': u'Schmo',
  u'middle initial': None,
  u'passport number': 22192,
  u'sex': u'M',
  u'weight': 100,
  u'weight_unit': u'lbs'},
 {u'date of birth': datetime.date(1967, 3, 20),
  u'expiration date': datetime.date(2017, 8, 16),
  u'first name': u'Sue',
  u'issue date': datetime.date(2002, 12, 23),
  u'issuing country': None,
  u'last name': u'Bobby',
  u'middle initial': None,
  u'passport number': None,
  u'sex': u'f',
  u'weight': 299,
  u'weight_unit': u'lbs'}]


def write_xlsx_file(data, output):
    '''
    :param data: a list of dict(s)
    :param output: uri for destination file (.xlsx) destructive updates that uri
    :return: True to indicate success
    '''

    header = [u'last name',
     u'middle initial',
     u'first name',
     u'date of birth',
     u'sex',
     u'weight',
     u'weight_unit',
     u'passport number',
     u'issuing country',
     u'issue date',
     u'expiration date']

    wb = load_workbook(output)
    ws = wb.active


    for item, row in zip(data, ws['D12':'N512']):
        for key, cell in zip(header, row):
            cell.value = item[key.lower()]

    wb.save(output)
    return True
