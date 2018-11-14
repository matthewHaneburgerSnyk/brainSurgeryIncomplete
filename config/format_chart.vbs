Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objExcel = CreateObject("Excel.Application")
Set objFile = objFSO.CreateTextFile("out.txt", True)
objFile.Write "Output to file test" & vbCrLf
objFile.Close