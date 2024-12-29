import unittest
import openpyxl as xl
import os

from code1 import (
    select_sheet,
    list_sheets,
    write_headers_to_sheet,
    check_if_exists_then_delete,
)

class TestFunction(unittest.TestCase):

    def setUp(self):
        # Create a temporary workbook for testing
        self.test_file = "test_workbook.xlsx"
        self.workbook = xl.Workbook()
        sheet1 = self.workbook.active
        sheet1.title = "Sheet1"
        sheet2 = self.workbook.create_sheet("Sheet2")
        sheet1.append(["Header1", "Header2", "Header3"])
        sheet2.append(["ColumnA", "ColumnB", "ColumnC"])
        self.workbook.save(self.test_file)

    def tearDown(self):
        # Clean up after tests
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_select_sheet(self):
        # Test that select_sheet prints the first few rows
        workbook = xl.load_workbook(self.test_file)
        with self.assertLogs(level='INFO') as log:
            select_sheet(workbook, "Sheet1")
        self.assertIn("Data from Sheet1:", log.output[0])