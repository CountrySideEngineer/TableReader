@echo off

rem Run test program to check spec of library, TableReader.SpecCheck.ClosedXML.dll.
rem This program reads TableReader_SpecCheck.xlsx excel file in TestData folder.
rem The sheet and table name to read is set to SHEET_NAME and TABLE_NAME variables.
rem And the number of test to read is set to TEST_COUNT.

setlocal

SET TEST_EXE=.\TableReader.SpecCheck.ExcelDataReader.exe
SET SHEET_NAME=Read_test_008
SET TABLE_NAME=TestTable_001
SET TEST_COUNT=10

%TEST_EXE% %SHEET_NAME% %TABLE_NAME% %TEST_COUNT%

endlocal
