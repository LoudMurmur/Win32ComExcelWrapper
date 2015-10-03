#!/usr/bin/env python
# --*-- encoding: iso-8859-1 --*--

import logging

from logging.handlers import RotatingFileHandler
from win32com import client

class Win32comExcelWrapper(object):
    def __init__(self):
        #avoid pop up blocking execution during automation
        self.displayAlerts = 0
        #Excel is visible
        self.visible = 1
        #you won't see what happens (faster)
        self.screenUpdating = False
        #clics on the Excel window have no effect
        #(set back to True before closing Excel)
        self.interactive = False
        self.initLogger()

    def initLogger(self):
        self.logger = logging.getLogger("Excel wrapper")
        #setting level to debug to see everything
        self.logger.setLevel(logging.DEBUG)
        #pretty self.logger.info(for logger
        formatter = logging.Formatter('%(asctime)s :: %(levelname)s :: %(message)s')
        #First logger writting on file, 1MB max
        file_handler = RotatingFileHandler('ExcelWrapper.log', 'a', 1000000, 1)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        #second logger writing over the console
        steam_handler = logging.StreamHandler()
        steam_handler.setLevel(logging.DEBUG)
        self.logger.addHandler(steam_handler)

    def openExcel(self):
        """ Lauch Excel and set it's configuration"""
        self.logger.info("Opening Excel")
        self.xl = client.Dispatch("Excel.Application")
        self.logger.info("Setting Excel configuration")
        self.xl.DisplayAlerts = self.displayAlerts
        self.xl.Visible = self.visible
        self.xl.ScreenUpdating = self.screenUpdating
        self.xl.Interactive = self.interactive

    def getWorkbook(self, workbook_absolute_path):
        """Open a workbook and returns it"""
        self.logger.info("Opening workbook %s" %workbook_absolute_path)
        return self.xl.Workbooks.Open(workbook_absolute_path)

    def closeWorkbookWithoutSaving(self, wb):
        """close workbook without saving
            -wb : a workbook object
        """
        self.logger.info("Closing workbook without saving")
        wb.Close()

    def saveWorkbook(self, wb):
        """Save the workbook passed in parameters
            -wb : a workbook object
        """
        self.logger.info("Saving workbook")
        wb.Save()

    def saveWorkbookAs(self, wb, filename):
        """Save the workbook passed in parameters with a new name
            -wb : a workbook object
            -filename : a string
        """
        self.logger.info("Saving workbook as %s" %filename)
        wb.SaveAs(filename)

    def closeExcel(self):
        """Put back Excel interativity online and close it"""
        self.logger.info("Re-activating Excel interactivity")
        self.xl.Interactive = True
        self.logger.info("Closing Excel")
        self.xl.Quit()

    class RangeCoordinate():
        """
        Class representing a Range Coordinate, typical use :

        data = ws.Range(
                         ws.Cells(coord.tline, coord.tcol),
                         ws.Cells(coord.bline, coord.bcol)
                        ).Value
        """
        def __init__(self, tline, tcol, bline, bcol):
            """Constructor
                -tline : top left line, int
                -tcol : top left column, int
                -bline : bottom right line, int
                -bcol : bottom right column, int
            """
            self.tline = tline
            self.tcol = tcol
            self.bline = bline
            self.bcol = bcol
