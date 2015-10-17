# Win32ComExcelWrapper
A nice way to do stuff with excel in python 2.7


############################
####    Prerequisite    ####
############################

1/Install python 2.7 :
https://www.python.org/download/releases/2.7/

create the PYTHON_HOME environment variable containing c:\Python27
Add ;%PYTHON_HOME%;%PYTHON_HOME%\Scripts to your PATH environment variable

2/Install pyWin32 (to get win32com)
http://sourceforge.net/projects/pywin32/?source=typ_redirect

3/Compile windows librairies (to get excel constants)
go to C:\Python27\Lib\site-package\client and run makepy.py
Select MICROSOFT Excel 14.0 Object library (1.7) and clic ok

4/Install nose for the unit tests
pip install nose


##########################
####    Vocabulary    ####
##########################

Workbook : an excel file
Worksheet : a sheet of the workbook


#####################
####    Usage    ####
#####################

There is a lot of method to read (read column with an int, etc).
However I advice to always read with : readAreaValuesExn(ws, excel_adress)

If you want to read colum J use "J:J" as the excel address, if you want to
read row 7 to 12 use "7:12" as excel adress, if you want to read a rectangle
use "C23:G45" for example, for a single cell use "C42", etc

I have written method to convert numerical coordinate to excel adress for
EVERY case.

those methode are the computeXXXXExcelAddress(xxxxx)

The other reading method will work fine, but I find the code prettier by using
only this method for reading.


##########################
####    Unit tests    ####
##########################

Every method is tested, run "nosetests -v" at the project root to launch them


#######################################
####    Additional informations    ####
#######################################

Tested for Excel 2010 on windows 7 and 8.1

