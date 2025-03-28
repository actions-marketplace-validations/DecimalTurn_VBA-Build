' Create a new Excel file using the Excel.Application COM object
'
'Print working directory
WScript.Echo "Current Directory: " & CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

Dim excel, wb

set excel = CreateObject("Excel.Application")

excel.Visible = true
set wb = excel.Workbooks.Add()

wb.Sheets(1).Cells(2, 1).Value = "Hello, World!"

wb.SaveAs "D:\a\Demo-Office-CLI\Demo-Office-CLI\example.xlsx"

wb.Close SaveChanges = false
excel.Quit

' Clean up
set wb = nothing
set excel = nothing

