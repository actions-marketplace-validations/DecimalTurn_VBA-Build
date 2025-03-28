' Create a new Excel file using the Excel.Application COM object
'
'Print working directory
WScript.Echo "Current Directory: " & CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

Dim excel, wb

set excel = CreateObject("Excel.Application")

excel.Visible = true
set wb = excel.Workbooks.Add()

wb.Sheets(1).Cells(2, 1).Value = "Hello, World!"

wb.SaveAs "C:\path\to\your\file.xlsx"



set excel = nothing

