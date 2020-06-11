import unittest
from lecture3_pdf import *
from lecture3_employee_payslip import *


dir = 'lecture3/'


class TestString(unittest.TestCase):

    def setUp(self):

        print("\n")
        print("201303024 신동규")

    def tearDown(self):
        print("Tear down")

    def test_read_pdf(self):
        reader, info = readPDF(f'{dir}/sample.pdf')
        print(f"Title: {info.title}")
        print(f"Author: {info.author}")
        print(f"Number of pages: {reader.getNumPages()}")
        print()

        print("Page 1:")
        page = reader.getPage(1)
        print(page.extractText())
        print()

        print("Outline")
        for heading in reader.getOutlines():
            if type(heading) is not list:
                print(dict(heading).get('/Title'))

    def test_write1_pdf(self):
        writePDF1(f'{dir}sample.pdf', f'{dir}myPDF.pdf')

    def test_write2_pdf(self):
        writePDF2(f'{dir}myPDF.pdf')

    def test_write3_pdf(self):
        writePDF3(dir)

    def test_write4_pdf(self):
        writePDF4(f'{dir}header_footer.pdf')

    def test_merger_pdf(self):
        filePath = mergerPDF(dir)

        reader, info = readPDF(filePath)
        print(info)

    def test_rotate_pdf(self):
        rotatePDF(f'{dir}sample.pdf')

    def test_encrypt_pdf(self):

        encryptPDF(f'{dir}sample.pdf', '1234')

    def test_employee_payslip(self):

        employee_data = [
            {
                'id': 111,
                'name': 'Donggyu Shin',
                'payment': 130000,
                'tax': 3000,
                'total': 17000,
                'birth': '941013'
            },
            {
                'id': 115,
                'name': 'John Smith',
                'payment': 12000,
                'tax': 2000,
                'total': 10000,
                'birth': '801010'
            }
        ]

        for emp in employee_data:
            filePath = generate_payslip(emp)


if __name__ == '__main__':
    unittest.main()
