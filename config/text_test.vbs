
Wscript.Echo "1"

    

	Set objExcel = CreateObject("Excel.Application") 
Wscript.Echo "2"
'view the excel program and file, set to false to hide the whole process
	objExcel.Visible = False
Wscript.Echo "3"
'open an excel file (make sure to change the location) .xls for 2003 or earlier
	Set objWorkbook = objExcel.Workbooks.Open("C:\Users\rl250177\Desktop\rubyBrain\graphing_tool\config\testChart.xlsx")
Wscript.Echo "4"
'set a cell value at row 3 column 5
	objExcel.Cells(3,5).Value = "new value"
Wscript.Echo "5"
'change a cell value
	objExcel.Cells(3,5).Value = "something different"
Wscript.Echo "6"	
'delete a cell value
	objExcel.Cells(3,5).Value = ""
Wscript.Echo "7"
'get a cell value and set it to a variable
	r3c5 = objExcel.Cells(3,5).Value
Wscript.Echo "8"
'save the existing excel file. use SaveAs to save it as something else
	objWorkbook.Save
    objExcel.Save
'close the workbook
	objWorkbook.Close 
Wscript.Echo "9"
'exit the excel program
	objExcel.Quit
Wscript.Echo "10"
'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing