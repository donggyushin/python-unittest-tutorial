import unittest
from lecture5_pptx import *

dir = 'lecture5/'


class TestString(unittest.TestCase):

    def setUp(self):
        print("신동규 201303024")
        print("\n")

    def tearDown(self):
        print("tear down")

    def test_read_ppt(self):
        prs = readPPT(f'{dir}myprofile.pptx')

        i = 1
        for slide in prs.slides:
            print('Slide: ', i)
            print('Layout: ', slide.slide_layout.name)
            for shape in slide.shapes:
                print('Shape: ', shape.shape_type)
            print()
            i += 1

        text_runs = []

        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text_runs.append(run.text)

        print('Text is: ', text_runs)

    def test_write_new_ppt(self):
        writeNewPPT(f'{dir}youPython.pptx')

    def test_write1_ppt(self):
        writePPT1(f'{dir}sample_ppt.pptx')

    def test_write2_ppt(self):
        writePPT2(f'{dir}two_content.pptx')

    def test_write3_ppt(self):
        writePPT3(f'{dir}textBox.pptx')

    def test_write_shapes_ppt(self):
        writeShapesPPT(f'{dir}shapes.pptx')

    def test_write_table_ppt(self):
        writeTablePPT(f'{dir}table.pptx')

    def test_challenge(self):
        challenge(f'{dir}challenge.pptx')
