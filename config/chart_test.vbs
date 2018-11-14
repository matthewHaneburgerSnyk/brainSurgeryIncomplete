Set obj = CreateObject("Excel.Application")  
obj.visible = True                                   
Set obj1 = obj.Workbooks.open("Z:/Users/rl250177/Desktop/rubyBrain/graphing_tool/config/testChart.xlsx")    
Set obj2 = obj1.Worksheets("Sheet1")   
'obj2.Rows("4:4").Delete         
obj2.Cells.ClearContents  
obj1.Save()                                  
obj1.Close                                             
obj.Quit                                                
Set obj1 = Nothing                               
Set obj2 = Nothing                             