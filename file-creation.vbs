
'Print working directory
WScript.Echo "Current Directory: " & CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

'Create a text file and save it in the current directory

Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("example.txt", True)
file.WriteLine("Hello, World!")
file.Close


