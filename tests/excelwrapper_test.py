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

    def test_getWorkbookGetWorksheet(self):
        """ test that get worksheet is working (so testing openWorkbook too)"""
        wrapper = Win32comExcelWrapper()
        wrapper.openExcel()
        ws_name = "FOO"
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        #if getWorksheet fail you get an exception here
        ws = wrapper.getWorksheet(wb, ws_name)
        self.assertEqual(ws_name, ws.Name)
        wrapper.closeExcel()

    def test_getWorksheetByNumber(self):
        """ test that get worksheet is working (so testing openWorkbook too)"""
        wrapper = Win32comExcelWrapper()
        wrapper.openExcel()
        ws_number = 1
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        #if getWorksheet fail you get an exception here
        ws = wrapper.getWorksheet(wb, ws_number)
        self.assertEqual(wb.Sheets(1).Name, ws.Name)
        wrapper.closeExcel()

    def test_copyWorksheet(self):
        wrapper = Win32comExcelWrapper()
        first_ws_name = "Feuil1"
        copy_ws_name = "DaCopy"

        #open workbook, write in it, then copy the sheet
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(first_ws_name)
        ws.Cells(1, 1).Value = "FOO"
        ws.Cells(3, 2).Value = "BAR"
        wrapper.copyWorksheet(wb, first_ws_name, copy_ws_name)
        wb.Sheets(copy_ws_name)

        #extract data from the copied sheet
        values = ws.Range(ws.Cells(1, 1), ws.Cells(3, 2)).Value
        wrapper.closeExcel()

        self.assertEqual((("FOO", None), (None, None), (None, "BAR")), values)

    def test_deleteWorksheetByName(self):
        wrapper = Win32comExcelWrapper()
        first_ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.deleteworksheet(wb, first_ws_name)
        try:
            wb.Sheets(first_ws_name)
            self.fail("%s has not been deleted" %first_ws_name)
        except:
            pass
        wrapper.closeExcel()

    def test_deleteWorksheetByPosition(self):
        wrapper = Win32comExcelWrapper()
        first_ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.deleteworksheet(wb, 1)
        try:
            wb.Sheets(first_ws_name)
            self.fail("%s has not been deleted" %first_ws_name)
        except:
            pass
        wrapper.closeExcel()

    def test_insertWorksheet(self):
        wrapper = Win32comExcelWrapper()
        ws_name = "freshLeaf"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.insertWorksheet(wb, 2, ws_name)
        try:
            _ = wb.Sheets(ws_name)
            self.fail("%s has not been inserted" %ws_name)
        except:
            pass
        self.assertEqual(ws_name, wb.Sheets(2).Name)
        wrapper.closeExcel()

    def test_moveWorksheetByName(self):
        wrapper = Win32comExcelWrapper()
        first_ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.moveWorksheet(wb, first_ws_name, 2)
        self.assertEqual(first_ws_name, wb.Sheets(2).Name)
        wrapper.closeExcel()

    def test_moveWorksheetBynumber(self):
        wrapper = Win32comExcelWrapper()
        first_ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.moveWorksheet(wb, 1, 2)
        self.assertEqual(first_ws_name, wb.Sheets(2).Name)
        wrapper.closeExcel()

    def test_renameWorksheetByName(self):
        wrapper = Win32comExcelWrapper()
        new_name = "FIRST!"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.renameworkSheet(wb, "Feuil1", "FIRST!")

        self.assertEqual(new_name, wb.Sheets(1).Name)
        wrapper.closeExcel()

    def test_renameWorksheetByPosition(self):
        wrapper = Win32comExcelWrapper()
        new_name = "FIRST!"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.renameworkSheet(wb, 1, "FIRST!")

        self.assertEqual(new_name, wb.Sheets(1).Name)
        wrapper.closeExcel()

    def test_hideSheetByName(self):
        wrapper = Win32comExcelWrapper()
        ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wb.Sheets(ws_name).Visible = True
        wrapper.hideSheet(wb, ws_name)
        self.assertEqual(False, wb.Sheets(ws_name).Visible)
        wrapper.closeExcel()

    def test_hideSheetByPosition(self):
        wrapper = Win32comExcelWrapper()
        ws_pos = 1

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wb.Sheets(ws_pos).Visible = True
        wrapper.hideSheet(wb, ws_pos)
        self.assertEqual(False, wb.Sheets("Feuil1").Visible) #hidden sheet has no position
        wrapper.closeExcel()

    def test_unhideSheet(self):
        wrapper = Win32comExcelWrapper()
        ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wb.Sheets(ws_name).Visible = False
        wrapper.unhideSheet(wb, ws_name)
        self.assertTrue(False != wb.Sheets(ws_name).Visible)
        wrapper.closeExcel()

    def _eraseSafely(self, path):
        if exists(path):
            os.remove(path)

    def _openWbAndExtractRange(self, wrapper, wb_name, ws_name, coord):
        """Open excel, read a range, close excel"""
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath(wb_name))
        ws = wb.Sheets(ws_name)
        values = ws.Range(
                          ws.Cells(coord.tline, coord.tcol),
                          ws.Cells(coord.bline, coord.bcol)
                         ).Value
        wrapper.closeExcel()
        return values

if __name__ == '__main__':
    unittest.main()
