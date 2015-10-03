#!/usr/bin/env python
# --*-- encoding: iso-8859-1 --*--

#/!\ A test method must always start by test_

import unittest
import util
import os
import time

from excelwrapper import Win32comExcelWrapper
from os.path import exists

class ExcelWrapperTest(unittest.TestCase):

    def test_openExcel(self):
        """Test that excel is configured with the right values"""
        wrapper = Win32comExcelWrapper()
        wrapper.openExcel()

        self.assertEqual(0, wrapper.xl.DisplayAlerts)
        self.assertEqual(1, wrapper.xl.Visible)
        self.assertEqual(False, wrapper.xl.ScreenUpdating)
        self.assertEqual(False, wrapper.xl.Interactive)

        wrapper.closeExcel()

    def test_getWorkbookGetWorksheet(self):
        """ test that get worksheet is working (so testing openWorkbook too)"""
        wrapper = Win32comExcelWrapper()
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        #if getWorksheet fail you get an exception here
        _ = wb.Sheets("FOO")
        wrapper.closeExcel()

    def test_saveWorkbookAs(self):
        wrapper = Win32comExcelWrapper()
        saved_wb_name = "saved.xlsx"
        first_ws_name = "Feuil1"

        #Write in workbook and save it
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(first_ws_name)
        ws.Cells(1, 1).Value = "FOO"
        ws.Cells(3, 2).Value = "BAR"
        wrapper.saveWorkbookAs(wb, util.getTestRessourcePath(saved_wb_name))
        wrapper.closeExcel()

        coord = Win32comExcelWrapper.RangeCoordinate(1, 1, 3, 2)

        #open target workbook and read expected value
        expected_values = self._openWbAndExtractRange(wrapper,
                                                      "testSaveAs_expected.xlsx",
                                                      first_ws_name,
                                                      coord)

        #open saved workbook and read written data
        writen_values = self._openWbAndExtractRange(wrapper,
                                            saved_wb_name,
                                            first_ws_name,
                                            coord)

        time.sleep(1) #let time for excel process to stop
        self._eraseSafely(util.getTestRessourcePath(saved_wb_name))

        self.assertEqual(expected_values, writen_values)

    def test_save(self):
        wrapper = Win32comExcelWrapper()
        blank_copy_name = "emptyWorkbookCopy.xlsx"
        first_ws_name = "Feuil1"

        #Write in workbook and save it elsewhere
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.saveWorkbookAs(wb, util.getTestRessourcePath(blank_copy_name))
        wrapper.closeExcel()

        #Open workbook, write insite, save it
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath(blank_copy_name))
        ws = wb.Sheets(first_ws_name)
        ws.Cells(1, 1).Value = "FOO"
        ws.Cells(3, 2).Value = "BAR"
        wrapper.saveWorkbookAs(wb, util.getTestRessourcePath(blank_copy_name))
        wrapper.closeExcel()

        coord = Win32comExcelWrapper.RangeCoordinate(1, 1, 3, 2)
        #Open target result, extract expected data
        expected_values = self._openWbAndExtractRange(wrapper,
                                                      "testSaveAs_expected.xlsx",
                                                      first_ws_name,
                                                      coord)
        #Open previously saved workbook, extract saved data
        writen_values = self._openWbAndExtractRange(wrapper,
                                                    blank_copy_name,
                                                    first_ws_name,
                                                    coord)

        time.sleep(1) #let time for excel process to stop
        self._eraseSafely(util.getTestRessourcePath(blank_copy_name))

        self.assertEqual(expected_values, writen_values)

    def _eraseSafely(self, path):
        if exists(path):
            os.remove(path)

    def _openWbAndExtractRange(self, wrapper, wb_name, ws_name, coord):
        """Open excel, read a range, close excel"""
        wrapper.openExcel()
        wb = wrapper.getWorkbook(util.getTestRessourcePath(wb_name))
        ws = wrapper.getWorksheet(wb, ws_name)
        values = ws.Range(
                          ws.Cells(coord.tline, coord.tcol),
                          ws.Cells(coord.bline, coord.bcol)
                         ).Value
        wrapper.closeExcel()
        return values

if __name__ == '__main__':
    unittest.main()
