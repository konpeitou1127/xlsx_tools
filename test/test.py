import unittest
from xlspy import xlspy
import openpyxl

class TestCreateDict(unittest.TestCase):

    def setUp(self):
        self.work_book = openpyxl.load_workbook("./test.xlsx")
        self.work_sheet = self.work_book.active

    def test_case1(self):

        result = xlspy.create_dict(self.work_sheet, "A1", "B3")

        correct_keys = ["リンゴ", "バナナ", "オレンジ"]
        correct_values = [100, 200, 300]

        for x, y in zip(result.keys(), correct_keys):
            assert x == y
        
        for x, y in zip(result.values(), correct_values):
            assert x == y

    def test_case2(self):
        result = xlspy.create_dict(self.work_sheet, "A1", "A1")

        correct_keys = ["リンゴ", "バナナ", "オレンジ"]
        correct_values = []

        for x, y in zip(result.keys(), correct_keys):
            assert x == y
        
        for x, y in zip(result.values(), correct_values):
            assert x == y


if __name__ == "__main__":
    unittest.main()
