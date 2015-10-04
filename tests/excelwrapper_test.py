#!/usr/bin/env python
# --*-- encoding: iso-8859-1 --*--

#/!\ A test method must always start by test_

import unittest
import util
import os
import time
import logmanager

from excelwrapper import Win32comExcelWrapper
from os.path import exists

class ExcelWrapperTest(unittest.TestCase):

    LOGGER = logmanager.getLogger("test wrapper")

    BEFORE = (
               (None, None, None, None, None, None, None, None),
               (None, None, None, None, None, None, None, None),
               (None, None, None, None, None, None, None, None),
               (None, u'a', u'b', u'c', u'd', u'e', None, None),
               (None, u'f', u'g', u'h', u'i', u'j', None, None),
               (None, u'k', u'l', u'm', u'n', u'o', None, None),
               (None, u'p', u'q', u'r', u's', u'r', None, None),
               (None, u'u', u'v', u'w', u'x', u'y', None, None),
               (None, u'z', u'aa', u'ab', u'ac', u'ad', None, None),
               (None, u'ae', u'af', u'ag', u'ah',u'ai', None, None),
               (None, None, None, None, None, None, None, None),
               (None, None, None, None, None, None, None, None),
               (None, None, None, None, None, None, None, None)
             )

    AREA = "A1:H13"

    def test_openExcel(self):
        """Test that excel is configured with the right values"""
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        wrapper.openExcel()

        self.assertEqual(0, wrapper.xl.DisplayAlerts)
        self.assertEqual(1, wrapper.xl.Visible)
        self.assertEqual(False, wrapper.xl.ScreenUpdating)
        self.assertEqual(False, wrapper.xl.Interactive)

        wrapper.closeExcel()

        wrapper.logger = self.LOGGER
    def test_saveWorkbookAs(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
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
        wrapper.logger = self.LOGGER
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
        wrapper.logger = self.LOGGER
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
        wrapper.logger = self.LOGGER
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
        wrapper.logger = self.LOGGER
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
        wrapper.logger = self.LOGGER
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
        wrapper.logger = self.LOGGER
        first_ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.moveWorksheet(wb, first_ws_name, 2)
        self.assertEqual(first_ws_name, wb.Sheets(2).Name)
        wrapper.closeExcel()

    def test_moveWorksheetBynumber(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        first_ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.moveWorksheet(wb, 1, 2)
        self.assertEqual(first_ws_name, wb.Sheets(2).Name)
        wrapper.closeExcel()

    def test_renameWorksheetByName(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        new_name = "FIRST!"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.renameworkSheet(wb, "Feuil1", "FIRST!")

        self.assertEqual(new_name, wb.Sheets(1).Name)
        wrapper.closeExcel()

    def test_renameWorksheetByPosition(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        new_name = "FIRST!"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wrapper.renameworkSheet(wb, 1, "FIRST!")

        self.assertEqual(new_name, wb.Sheets(1).Name)
        wrapper.closeExcel()

    def test_hideSheetByName(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wb.Sheets(ws_name).Visible = True
        wrapper.hideSheet(wb, ws_name)
        self.assertEqual(False, wb.Sheets(ws_name).Visible)
        wrapper.closeExcel()

    def test_hideSheetByPosition(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_pos = 1

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wb.Sheets(ws_pos).Visible = True
        wrapper.hideSheet(wb, ws_pos)
        self.assertEqual(False, wb.Sheets("Feuil1").Visible) #hidden sheet has no position
        wrapper.closeExcel()

    def test_unhideSheet(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        wb.Sheets(ws_name).Visible = False
        wrapper.unhideSheet(wb, ws_name)
        self.assertTrue(False != wb.Sheets(ws_name).Visible)
        wrapper.closeExcel()

    def test_writeCellValue(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        wrapper.writeCellValue(ws, 42, 42, "H2G2")
        wrapper.writeCellValue(ws, 12, 34, 56)
        cell42_value = ws.Cells(42, 42).Value
        cell1234_value = ws.Cells(12, 34).Value
        wrapper.closeExcel()

        self.assertEqual("H2G2", cell42_value)
        self.assertEqual(56, cell1234_value)

    def test_writeCellFormula(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"
        formula = "=SUM(R5C4:RC[-18])"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        wrapper.writeCellFormula(ws, 42, 42, formula)
        cell42_value = ws.Cells(42, 42).FormulaR1C1
        wrapper.closeExcel()

        self.assertEqual(formula, cell42_value)

    def test_writeCell(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"
        formula = "=SUM(R5C4:RC[-18])"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        wrapper.writeCell(ws, 42, 42, "H2G2")
        wrapper.writeCell(ws, 12, 34, formula)
        cell42_value = ws.Cells(42, 42).Value
        cell1234_formula = ws.Cells(12, 34).FormulaR1C1
        wrapper.closeExcel()

        self.assertEqual("H2G2", cell42_value)
        self.assertEqual(formula, cell1234_formula)

    def test_writeAreaCellByCell(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"
        formula = "=SUM(R5C4:RC[-1])"
        data = (
               ("this", "is", "now", "written"),
               (None, "in", "Excel", None),
               (None, None, None, formula)
               )

        expected_value_range = (
                                (u"this", u"is", u"now", u"written"),
                                (None, u"in", u"Excel", None),
                                (None, None, None, 0.0)
                               )

        expected_formulas_range = (
                                   (u"this", u"is", u"now", u"written"),
                                   (u"", u"in", u"Excel", u""),
                                   (u"", u"", u"", unicode(formula))
                                  )
        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        wrapper.writeAreaCellByCell(ws, 5, 8, data)

        written_data_values = ws.Range(
                                       ws.Cells(5, 8),
                                       ws.Cells(7, 11)
                                       ).Value
        written_data_formulas = ws.Range(
                                         ws.Cells(5, 8),
                                         ws.Cells(7, 11)
                                         ).FormulaR1C1

        wrapper.closeExcel()

        self.assertEqual(expected_value_range, written_data_values)
        self.assertEqual(expected_formulas_range, written_data_formulas)

    def test_writeAreaInOneCall(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "Feuil1"
        formula = "=SUM(R5C4:RC[-5])"
        data = (
               ("this", "is", "now", "written"),
               (None, "in", "Excel", None),
               (None, "foo", None, formula)
               )

        expected_value_range = (
                                (u"this", u"is", u"now", u"written"),
                                (None, u"in", u"Excel", None),
                                (None, "foo", None, 0.0)
                               )

        expected_formulas_range = (
                                   (u"this", u"is", u"now", u"written"),
                                   (u"", u"in", u"Excel", u""),
                                   (u"", u"foo", u"", unicode(formula))
                                  )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("emptyWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        wrapper.writeAreaInOneCall(ws, 5, 8, data)

        written_data_values = ws.Range(
                                       ws.Cells(5, 8),
                                       ws.Cells(7, 11)
                                       ).Value
        written_data_formulas = ws.Range(
                                         ws.Cells(5, 8),
                                         ws.Cells(7, 11)
                                         ).FormulaR1C1

        wrapper.closeExcel()

        self.assertEqual(expected_value_range, written_data_values)
        self.assertEqual(expected_formulas_range, written_data_formulas)

    def test_readCellValue(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        value_3203 = wrapper.readCellValue(ws, 32, 3)
        wrapper.closeExcel()

        self.assertEqual(u"yolo", value_3203)

    def test_readCellValueExn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        value_C32 = wrapper.readCellValueExn(ws, "C32")
        value_C32_2 = wrapper.readCellValueExn(ws, "$C$32")
        wrapper.closeExcel()

        self.assertEqual(u"yolo", value_C32)
        self.assertEqual(u"yolo", value_C32_2)

    def test_readAreaValues(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        expected_data = ((None, None, None, u'r', None),
                         (u'c', u'd', u'e', u'f', u'g'),
                         (None, u'e', u'f', u'g', u'h'),
                         (None, None, None, None, None),
                         (u'f', u'g', u'h', u'i', None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, u'e', u'f', u'g'))

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        coord = Win32comExcelWrapper.RangeCoordinate(1, 3, 13, 7)
        data = wrapper.readAreaValues(ws, coord)
        wrapper.closeExcel()

        self.assertEqual(expected_data, data)

    def test_readAreaValuesExn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        expected_data = ((None, None, None, u'r', None),
                         (u'c', u'd', u'e', u'f', u'g'),
                         (None, u'e', u'f', u'g', u'h'),
                         (None, None, None, None, None),
                         (u'f', u'g', u'h', u'i', None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, None, None, None),
                         (None, None, u'e', u'f', u'g'))

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        data = wrapper.readAreaValuesExn(ws, "C1:G13")
        data2 = wrapper.readAreaValuesExn(ws, "$C$1:$G$13")
        wrapper.closeExcel()

        self.assertEqual(expected_data, data)
        self.assertEqual(expected_data, data2)

    def test_computeCellExcelAddress(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        adr = wrapper.computeCellExcelAddress(ws, 7, 3)
        wrapper.closeExcel()

        self.assertEqual("$C$7", adr)

    def test_ComputeColumnExcelAddress(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        adr = wrapper.ComputeColumnExcelAddress(ws, 10)
        wrapper.closeExcel()

        self.assertEqual("$J:$J", adr)

    def test_computeColumnsExcelAddress(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        adr = wrapper.computeColumnsExcelAddress(ws, 5, 9)
        wrapper.closeExcel()

        self.assertEqual("$E:$I", adr)

    def test_computeRowExcelAddress(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        adr = wrapper.computeRowExcelAddress(ws, 7)
        wrapper.closeExcel()

        self.assertEqual("$7:$7", adr)

    def test_computeRowsExcelAddress(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        adr = wrapper.computeRowsExcelAddress(ws, 2, 10)
        wrapper.closeExcel()

        self.assertEqual("$2:$10", adr)

    def test_computeAreaExcelAddress(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)
        coord = Win32comExcelWrapper.RangeCoordinate(1, 3, 13, 7)
        adr = wrapper.computeAreaExcelAddress(ws, coord)

        wrapper.closeExcel()

        self.assertEqual("$C$1:$G$13", adr)

    def test_computeColumnLastLine(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"
        exp_results = [2, 5, 32, 5, 13, 13, 13, 13, 1]

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        results = []
        results.append(wrapper.computeColumnLastLine(ws, 1))
        results.append(wrapper.computeColumnLastLine(ws, 2))
        results.append(wrapper.computeColumnLastLine(ws, 3))
        results.append(wrapper.computeColumnLastLine(ws, 4))
        results.append(wrapper.computeColumnLastLine(ws, 5))
        results.append(wrapper.computeColumnLastLine(ws, 6))
        results.append(wrapper.computeColumnLastLine(ws, 7))
        results.append(wrapper.computeColumnLastLine(ws, 8))
        results.append(wrapper.computeColumnLastLine(ws, 20))

        wrapper.closeExcel()

        self.assertEqual(exp_results, results)

    def test_computeLastColumn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        last_col = wrapper.computeLastColumn(ws)
        wrapper.closeExcel()

        self.assertEqual(18, last_col)

    def test_computeAreaFromData(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        ws_name = "FOO"

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("filledWorkbook.xlsx"))
        ws = wb.Sheets(ws_name)

        data = [
                [0, 0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0, 0],
                [0, 0, 0, 0, 0, 0]
               ]

        adr = wrapper.computeAreaAddressFromData(ws, 3, 3, data)
        wrapper.closeExcel()

        self.assertEqual("C3:H6", adr)

    def test_insertEmptyRow(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', u'b', u'c', u'd', u'e', None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'f', u'g', u'h', u'i', u'j', None, None),
                          (None, u'k', u'l', u'm', u'n', u'o', None, None),
                          (None, u'p', u'q', u'r', u's', u'r', None, None),
                          (None, u'u', u'v', u'w', u'x', u'y', None, None),
                          (None, u'z', u'aa', u'ab', u'ac', u'ad', None, None),
                          (None, u'ae', u'af', u'ag', u'ah', u'ai', None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                          )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.insertEmptyRow(ws, 5)
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_insertEmptyRowExn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', u'b', u'c', u'd', u'e', None, None),
                          (None, u'f', u'g', u'h', u'i', u'j', None, None),
                          (None, u'k', u'l', u'm', u'n', u'o', None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'p', u'q', u'r', u's', u'r', None, None),
                          (None, u'u', u'v', u'w', u'x', u'y', None, None),
                          (None, u'z', u'aa', u'ab', u'ac', u'ad', None, None),
                          (None, u'ae', u'af', u'ag', u'ah', u'ai', None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                          )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.insertEmptyRow(ws, "7:7")
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_insertEmptyColumn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', None, u'b', u'c', u'd', u'e', None),
                          (None, u'f', None, u'g', u'h', u'i', u'j', None),
                          (None, u'k', None, u'l', u'm', u'n', u'o', None),
                          (None, u'p', None, u'q', u'r', u's', u'r', None),
                          (None, u'u', None, u'v', u'w', u'x', u'y', None),
                          (None, u'z', None, u'aa', u'ab', u'ac', u'ad', None),
                          (None, u'ae', None, u'af', u'ag', u'ah', u'ai', None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                         )


        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.insertEmptyColumn(ws, 3)
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_insertEmptyColumnExn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', u'b', u'c', None, u'd', u'e', None),
                          (None, u'f', u'g', u'h', None, u'i', u'j', None),
                          (None, u'k', u'l', u'm', None, u'n', u'o', None),
                          (None, u'p', u'q', u'r', None, u's', u'r', None),
                          (None, u'u', u'v', u'w', None, u'x', u'y', None),
                          (None, u'z', u'aa', u'ab', None, u'ac', u'ad', None),
                          (None, u'ae', u'af', u'ag', None, u'ah', u'ai', None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                         )


        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.insertEmptyColumn(ws, "E:E")
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_deleteOneRow(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', u'b', u'c', u'd', u'e', None, None),
                          (None, u'f', u'g', u'h', u'i', u'j', None, None),
                          (None, u'k', u'l', u'm', u'n', u'o', None, None),
                          (None, u'u', u'v', u'w', u'x', u'y', None, None),
                          (None, u'z', u'aa', u'ab', u'ac', u'ad', None, None),
                          (None, u'ae', u'af', u'ag', u'ah', u'ai', None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                         )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.deleteRow(ws, 7)
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_deleteSeveralRows(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                           (None, None, None, None, None, None, None, None),
                           (None, None, None, None, None, None, None, None),
                           (None, None, None, None, None, None, None, None),
                           (None, u'a', u'b', u'c', u'd', u'e', None, None),
                           (None, u'u', u'v', u'w', u'x', u'y', None, None),
                           (None, u'z', u'aa', u'ab', u'ac', u'ad', None, None),
                           (None, u'ae', u'af', u'ag', u'ah',u'ai', None, None),
                           (None, None, None, None, None, None, None, None),
                           (None, None, None, None, None, None, None, None),
                           (None, None, None, None, None, None, None, None),
                           (None, None, None, None, None, None, None, None),
                           (None, None, None, None, None, None, u'test test', u'chuuuchuu'),
                           (None, None, None, None, None, None, None, None)
                         )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.deleteRow(ws, "5:7")
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_deleteOneColumn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', u'b', u'd', u'e', None, None, None),
                          (None, u'f', u'g', u'i', u'j', None, None, None),
                          (None, u'k', u'l', u'n', u'o', None, None, None),
                          (None, u'p', u'q', u's', u'r', None, None, None),
                          (None, u'u', u'v', u'x', u'y', None, None, None),
                          (None, u'z', u'aa', u'ac', u'ad', None, None, None),
                          (None, u'ae', u'af', u'ah', u'ai', None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                         )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.deleteColumn(ws, 4)
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_deleteSeveralColumn(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        expected_after = (
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, u'a', u'e', None, None, None, None, 1.0),
                          (None, u'f', u'j', None, None, None, None, None),
                          (None, u'k', u'o', None, None, None, None, 5.0),
                          (None, u'p', u'r', None, None, None, None, 5.0),
                          (None, u'u', u'y', None, None, None, None, 5.0),
                          (None, u'z', u'ad', None, None, None, None, 5.0),
                          (None, u'ae', u'ai', None, None, None,None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None),
                          (None, None, None, None, None, None, None, None)
                         )

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        data_before = ws.Range(self.AREA).Value
        wrapper.deleteColumn(ws, "C:E")
        data_after = ws.Range(self.AREA).Value

        wrapper.closeExcel()

        self.assertEqual(data_before, self.BEFORE)
        self.assertEqual(expected_after, data_after)

    def test_clearCell(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        white_color = 7405514.0

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        wrapper.clearCell(ws, 10, 2)
        wrapper.clearCell(ws, 6, 11)
        wrapper.clearCell(ws, 14, 7)
        wrapper.clearCell(ws, 20, 7)

        cellB10_value = ws.Cells(10, 2).Value
        cellK6_formula = ws.Cells(6, 11).FormulaR1C1
        cellG14_fontColor = ws.Cells(14, 7).Font.Color
        cellG20_color = ws.Cells(20, 7).Interior.Color
        wrapper.closeExcel()

        self.assertEqual(None, cellB10_value)
        self.assertEqual(u'', cellK6_formula)
        self.assertEqual(0.0, cellG14_fontColor)
        self.assertEqual(white_color, cellG20_color)

    def test_clearArea(self):
        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER
        white_color = 7405514.0

        wrapper.openExcel()
        wb = wrapper.xl.Workbooks.Open(util.getTestRessourcePath("insertDelete.xlsx"))
        ws = wb.Sheets(1)

        wrapper.clearArea(ws, "B6:M22")

        cellB10_value = ws.Cells(10, 2).Value
        cellK6_formula = ws.Cells(6, 11).FormulaR1C1
        cellG14_fontColor = ws.Cells(14, 7).Font.Color
        cellG20_color = ws.Cells(20, 7).Interior.Color
        wrapper.closeExcel()

        self.assertEqual(None, cellB10_value)
        self.assertEqual(u'', cellK6_formula)
        self.assertEqual(0.0, cellG14_fontColor)
        self.assertEqual(white_color, cellG20_color)

    def test_copyPaste(self):

        wrapper = Win32comExcelWrapper()
        wrapper.logger = self.LOGGER

        wrapper.openExcel()

        def copyPasteColumns():

            wrapper.logger.info("#### Testing copyPasteColumns ####")
            wb = wrapper.getWorkbook(util.getTestRessourcePath("copyPaste.xls"))

            def copyPasteColumnsAllNumerical():
                ws = wb.Sheets(1)
                first_col = wrapper.readAreaValuesExn(ws, "A1:A25")
                wrapper.copyPasteColumns(ws, ws, 1, 4)
                pasted_col = wrapper.readAreaValuesExn(ws, "D1:D25")
                self.assertEqual(first_col, pasted_col)

            def copyPastecolumnsSourceNumerical():
                ws = wb.Sheets(2)
                first_col = wrapper.readAreaValuesExn(ws, "A1:A25")
                wrapper.copyPasteColumns(ws, ws, 1, 'D:D')
                pasted_col = wrapper.readAreaValuesExn(ws, "D1:D25")
                self.assertEqual(first_col, pasted_col)

            def copyPasteColumnsDestNumerical():
                ws = wb.Sheets(3)
                first_twocol = wrapper.readAreaValuesExn(ws, "A1:B25")
                wrapper.copyPasteColumns(ws, ws, 'A:B', 4)
                pasted_twocol = wrapper.readAreaValuesExn(ws, "D1:E25")
                self.assertEqual(first_twocol, pasted_twocol)

            def copyPasteColumnsAllExcelAdress():
                ws = wb.Sheets(4)
                first_twocol = wrapper.readAreaValuesExn(ws, "A1:B25")
                wrapper.copyPasteColumns(ws, ws, 'A:B', 'D:D')
                pasted_twocol = wrapper.readAreaValuesExn(ws, "D1:E25")
                self.assertEqual(first_twocol, pasted_twocol)

            def copyPasteColumnsCutMode():
                ws = wb.Sheets(5)
                first_twocol = wrapper.readAreaValuesExn(ws, "A1:B25")
                empty_twocol = wrapper.readAreaValuesExn(ws, "C1:D25")
                wrapper.copyPasteColumns(ws, ws, 'A:B', 'D:D', True)
                pasted_twocol = wrapper.readAreaValuesExn(ws, "D1:E25")
                first_twocol_aftercut = wrapper.readAreaValuesExn(ws, "A1:B25")
                self.assertEqual(first_twocol, pasted_twocol)
                self.assertEqual(empty_twocol, first_twocol_aftercut)

            copyPasteColumnsAllNumerical()
            copyPastecolumnsSourceNumerical()
            copyPasteColumnsDestNumerical()
            copyPasteColumnsAllExcelAdress()
            copyPasteColumnsCutMode()

            wrapper.closeWorkbookWithoutSaving(wb)

        def copyPasteRows():

            wrapper.logger.info("#### Testing copyPasteColumns ####")
            wb = wrapper.getWorkbook(util.getTestRessourcePath("copyPaste.xls"))

            def copyPasteRowsAllNumerical():
                ws = wb.Sheets(1)
                first_row = wrapper.readAreaValuesExn(ws, "A1:B1")
                wrapper.copyPasteRows(ws, ws, 1, 40)
                pasted_row = wrapper.readAreaValuesExn(ws, "A40:B40")
                self.assertEqual(first_row, pasted_row)

            def copyPasteRowsSourceNumerical():
                ws = wb.Sheets(2)
                first_row = wrapper.readAreaValuesExn(ws, "A1:B1")
                wrapper.copyPasteRows(ws, ws, 1, '40:40')
                pasted_row = wrapper.readAreaValuesExn(ws, "A40:B40")
                self.assertEqual(first_row, pasted_row)

            def copyPasteRowsDestNumerical():
                ws = wb.Sheets(3)
                first_3row = wrapper.readAreaValuesExn(ws, "A1:B3")
                wrapper.copyPasteRows(ws, ws, '1:3', 40)
                pasted_3row = wrapper.readAreaValuesExn(ws, "A40:B42")
                self.assertEqual(first_3row, pasted_3row)

            def copyPasteRowsAllExcelAdress():
                ws = wb.Sheets(4)
                first_3row = wrapper.readAreaValuesExn(ws, "A1:B3")
                wrapper.copyPasteRows(ws, ws, '1:3', '40:40')
                pasted_3row = wrapper.readAreaValuesExn(ws, "A40:B42")
                self.assertEqual(first_3row, pasted_3row)

            def copyPasteRowsCutMode():
                ws = wb.Sheets(5)
                first_3row = wrapper.readAreaValuesExn(ws, "A1:B3")
                empty_3row = wrapper.readAreaValuesExn(ws, "A26:B28")
                wrapper.copyPasteRows(ws, ws, '1:3', '40:40', True)
                pasted_3row = wrapper.readAreaValuesExn(ws, "A40:B42")
                first_3row_aftercut = wrapper.readAreaValuesExn(ws, "A1:B3")
                self.assertEqual(first_3row, pasted_3row)
                self.assertEqual(empty_3row, first_3row_aftercut)

            copyPasteRowsAllNumerical()
            copyPasteRowsSourceNumerical()
            copyPasteRowsDestNumerical()
            copyPasteRowsAllExcelAdress()
            copyPasteRowsCutMode()

            wrapper.closeWorkbookWithoutSaving(wb)

        def copyAreaRows():

            wrapper.logger.info("#### Testing copyPasteColumns ####")
            wb = wrapper.getWorkbook(util.getTestRessourcePath("copyPaste.xls"))

            def copyPasteArea():
                ws = wb.Sheets(1)
                area_values = wrapper.readAreaValuesExn(ws, "A1:B10")
                wrapper.copyPasteArea(ws, ws, "A1:B10", "G40")
                pasted_area = wrapper.readAreaValuesExn(ws, "G40:H49")
                self.assertEqual(area_values, pasted_area)

            def copyPasteAreaCutMode():
                ws = wb.Sheets(2)
                area_values_before = wrapper.readAreaValuesExn(ws, "A1:B10")
                empty_area_values = wrapper.readAreaValuesExn(ws, "C1:D10")
                wrapper.copyPasteArea(ws, ws, "A1:B10", "G40", True)
                pasted_area_values = wrapper.readAreaValuesExn(ws, "G40:H49")
                area_aftercut = wrapper.readAreaValuesExn(ws, "A1:B10")
                self.assertEqual(area_values_before, pasted_area_values)
                self.assertEqual(empty_area_values, area_aftercut)

            def pasteAreaOnOtherWorksheet():
                ws = wb.Sheets(1)
                dst_ws = wb.Sheets(6)
                area_values = wrapper.readAreaValuesExn(ws, "A1:B10")
                wrapper.copyPasteArea(ws, dst_ws, "A1:B10", "G40")
                pasted_area = wrapper.readAreaValuesExn(dst_ws, "G40:H49")
                self.assertEqual(area_values, pasted_area)

            copyPasteArea()
            copyPasteAreaCutMode()
            pasteAreaOnOtherWorksheet()

            wrapper.closeWorkbookWithoutSaving(wb)

        copyPasteColumns()
        copyPasteRows()
        copyAreaRows()

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
