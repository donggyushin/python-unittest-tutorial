import unittest
from lecture4_docx import *

dir = 'lecture4/'


class TestString(unittest.TestCase):

    def setUp(self):
        print('\n')
        print('201303024 신동규')

    def tearDown(self):
        print("tear down")

    def test_read_document(self):
        doc, prop = readDocument(f'{dir}sample.docx')

        print('Title: %s' % doc.paragraphs[0].text)
        print('Author: %s' % prop.author)
        print('Created: %s' % prop.created)
        print('Revision: %s' % prop.revision)
        print('')
