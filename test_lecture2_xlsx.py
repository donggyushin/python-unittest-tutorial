import unittest
from lecture2_xlsx import *

dir = 'lecture2/'


class TestString(unittest.TestCase):

    def setUp(self):
        print('\n')

    def tearDown(self):
        print()

    def test_read_xlsx(self):
        book, sheet = read_xlsx(dir + 'myxlsx.xlsx', 'People')
        print('Sheets: ', book.sheetnames)
        print('Sheet: ', sheet.title)
        print('First cell: ', sheet['A1'].value)
        print('Other cell: ', sheet.cell(row=3, column=2).value)

        fm = '%-7s'
        for r in sheet.rows:
            print(f"{fm}: %s, %s" % (r[0].value, r[1].value, r[2].value))

    def test_write_xlsx(self):
        data = (
            ['Rent', 1000],
            ['Gas', 100],
            ['Food', 300],
            ['Gym', 50]
        )

        write_xlsx(dir + 'expenses.xlsx', 'expenses', data)

    def test_write_format_xlsx(self):
        data = (
            ['Rent', 1000],
            ['Gas', 100],
            ['Food', 300],
            ['Gym', 50]
        )

        cells = 'B1:B4'
        fm = {
            'bg_color': 'blue'
        }

        condition = {
            'type': 'cell',
            'criteria': '>=',
            'value': 150
        }

        white_format_xlsx(dir + 'exprenses.xlsx', 'expenses',
                          data, cells, fm, condition)

    def test_write_chart_xlsx(self):
        data = [10, 40, 50, 20, 10, 50]
        cells = 'C1'
        chart = {
            'type': 'line'
        }
        series = {
            'values': 'Chart!$A$1:$A$6'
        }
        write_chart_xlsx(dir + 'chart_line.xlsx', 'chart',
                         data, cells, chart, series)

    def test_financial_analysis(self):
        financial_analysis(dir + 'financial_analysis.xlsx')


if __name__ == '__main__':
    unittest.main()
