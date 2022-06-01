Set objExcel=CreateObject("Excel.Application")
Set objWorkbook=objExcel.Workbooks.Open("C:\Users\praprade\Desktop\OrderChecker\Orderability Status Check Input Form_V1.0.xlsm")

objExcel.Application.Run "Mail_workbook_Outlook"
objExcel.ActiveWorkbook.Close

objExcel.Application.Quit
WScript.Quit