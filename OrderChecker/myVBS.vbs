Set objExcel=CreateObject("Excel.Application")
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
objScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objWorkbook=objExcel.Workbooks.Open(objScriptDir & "\Orderability Status Check Input Form_V1.0.xlsm")

objExcel.Application.Run "Mail_workbook_Outlook"
objExcel.ActiveWorkbook.Close

objExcel.Application.Quit
WScript.Quit