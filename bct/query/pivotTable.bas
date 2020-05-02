REM  *****  BASIC  *****

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
  Dim openedDoc As Object
  Dim oRange As Object
  Dim totalTime As Long
  Dim Max As Long
  Dim Min As Long
   
  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials
   
  'RELATIVE_PATH ---> assign directory path here
  'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
  FILE_PATH = RELATIVE_PATH & "/" & FILE_PREFIX & "/" & (rowSize) & ".ods"
  url = ConvertToURL(FILE_PATH)
  
  openedDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open document
  DataCellRange = openedDoc.sheets(0).getCellRangeByName("A1:P" & (rowSize))
  RangeAddress = DataCellRange.RangeAddress

  For j = 0 To 9
     lTick = GetSystemTicks()
     Tables = openedDoc.sheets(0).DataPilotTables()   'Tables has all the DataPilot Tables in the Active Sheet
		
     'This part of the code just removes the table if it already exists. Prevents error from running the code several times
     If Tables.hasByName("NewDataPilot") THEN 
      Tables.removeByName("NewDataPilot")
     End If
		
     Descriptor = Tables.createDataPilotDescriptor()    'Descriptor contains the description of a DataPilot Table
     Descriptor.ShowFilterButton = False           'Don't show the Filter Button
     Descriptor.setSourceRange(RangeAddress)   'RangeAddress is defined above to cover A1:C10     
     Fields = Descriptor.getDataPilotFields      
     dimension = Fields.getByIndex(1)   'The first column of the data range has index 0
		
     'Set the Enum DataPilotFieldOrientation from com.sun.star.sheet.DataPilotField
     dimension.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.ROW                  
		
     measure = Fields.getByIndex(9)
     measure.Orientation = com.sun.star.sheet.DataPilotFieldOrientation.DATA
     measure.Function = com.sun.star.sheet.GeneralFunction.SUM
     Descriptor.RowGrand = "FALSE"   'Turn off the Total line of the Table
     Cell = openedDoc.sheets(0).getCellrangeByName("AA1")
     Tables.insertNewByName("NewDataPilot", Cell.CellAddress, Descriptor)
     lTick = (GetSystemTicks() - lTick)

     Tables.removeByName("NewDataPilot") 'delete pivot table
     totalTime = totalTime + lTick
     
     If lTick > Max Then
      Max = lTick
     End If
     
     If lTick < Min Then
      Min = lTick
     End If
  Next j
  
  openedDoc.dispose 'close document

  totalTime = totalTime - Max - Min 'remove outliers
  
  'write results back to oDoc  
  oSheet.getCellByPosition(0,rowIndex).String = rowCount
  oSheet.getCellByPosition(1, rowIndex).String = totalTime/8
End Sub

'Runs experiments on all spreadsheets specified by  [minRows, maxRows] with stepSize increments.
'This is the main function to be called for running the experiment.

Sub main
  Dim oDoc As Object
  Dim oSheet As Object
  Dim minRows as Long
  Dim maxRows as Long
  Dim stepSize as Long
  Dim rowIndex as Long
  
  oDoc = ThisComponent ' the file where the results are written
  oSheet = oDoc.Sheets(0) 'get first sheet
  minRows = 10000 `min row size
  maxRows = 50000 `max row size
  stepSize = 10000 `increment row sizes by 10k
  
  'add headers to the current file where results will be written
  oSheet.getCellByPosition(0,0).String = "Import Size" 
  oSheet.getCellByPosition(1, 0).String = "Time (ms)" 

  rowIndex = 1 'row id where the current result will be written
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i)
    rowIndex = rowIndex + 1   
  Next i
End Sub
