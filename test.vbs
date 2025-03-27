' Create a new Excel file using the Excel.Application COM object
'
Dim excel

set excel = CreateObject("Excel.Application")

excel.Visible = true
excel.Workbooks.Add()

excel.Cells(2, 1).Value = "Hello, World!"

set excel = nothing

