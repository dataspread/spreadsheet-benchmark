REM  *****  BASIC  *****

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
    Dim openedDoc As Object
    Dim oRange As Object
 
   
    Dim totalTime As Long
    Dim Max As Long
    Dim Min As Long
   
   
    Max = -1
    Min = 1000000
    totalTime = 0
   
    'RELATIVE_PATH ---> assign directory path here
    'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
    FILE_PATH = RELATIVE_PATH & "/" & FILE_PREFIX & "/" & (rowSize) & ".ods"
    url = ConvertToURL(FILE_PATH)
    
    openedDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open it
      

    For j = 0 To 9
       lTick = GetSystemTicks()
       oController = ThisComponent.CurrentController
       oSheetObj = oController.ActiveSheet
       DataCellRange = oSheetObj.getCellRangeByName("A1:P" & (rowCount))
       RangeAddress = DataCellRange.RangeAddress
       Tables = oSheetObj.DataPilotTables()   'Tables has all the DataPilot Tables in the Active Sheet
		
       'This part of the code just removes the table if it already exists. Prevents error from running the code several times
       If Tables.hasByName("NewDataPilot") THEN 
          Tables.removeByName("NewDataPilot")
       End If
		
       Descriptor = Tables.createDataPilotDescriptor()      'Descriptor contains the description of a DataPilot Table
       Descriptor.ShowFilterButton = False                   'Don't show the Filter Button
       Descriptor.setSourceRange(RangeAddress)     'RangeAddress is defined above to cover A1:C10       
       Fields = Descriptor.getDataPilotFields            
       dimension = Fields.getByIndex(1)   'The first column of the data range has index 0
		
       'Set the Enum DataPilotFieldOrientation from com.sun.star.sheet.DataPilotField
       dimension.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.ROW                                    
		
       measure = Fields.getByIndex(9)
       measure.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.DATA
       measure.Function = com.sun.star.sheet.GeneralFunction.SUM
       Descriptor.RowGrand = "FALSE"   'Turn off the Total line of the Table
       Cell = oSheetObj.getCellrangeByName("AA1")
       Tables.insertNewByName("NewDataPilot", Cell.CellAddress, Descriptor)
       lTick = (GetSystemTicks() - lTick)
       Tables.removeByName("NewDataPilot")
       totalTime = totalTime + lTick
         
       If lTick > Max Then
          Max = lTick
       End If
         
       If lTick < Min Then
          Min = lTick
       End If
    Next j
    
    openedDoc.dispose 'close it

    totalTime = totalTime - Max - Min
    
    oSheet.getCellByPosition(0,rowIndex).String = rowCount
    oSheet.getCellByPosition(1, rowIndex).String = totalTime/8

End Sub

Sub main
    Dim oDoc As Object
    Dim oSheet As Object
    Dim j as Long

    oDoc = ThisComponent ' the file where the results are written

    oSheet = oDoc.Sheets(0) 'get Sheet1
   
    oSheet.getCellByPosition(0,0).String = "Import Size"
   
    oSheet.getCellByPosition(1, 0).String = "Time (ms)"

    rowIndex = 1 'row id where the current result will be written
   
    For i = 10000 to 500001 Step 10000
        calculateRunTime(oDoc,oSheet,rowIndex,i)
        rowIndex = rowIndex + 1   
    Next i
End Sub
