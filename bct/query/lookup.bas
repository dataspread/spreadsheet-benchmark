REM  *****  BASIC  *****

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long, is_sorted as Long)
  Dim oCellRange As Object
  Dim oSvc as variant
  Dim oArg as variant
  Dim OldDoc As Object
  Dim OldSheet As Object
  Dim lTick As Long
   
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
  OldDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open document

  oSvc = createUnoService( "com.sun.star.sheet.FunctionAccess")
  lookupIndex = 200000
  returnIndex = 3
  oCellRange = OldDoc.sheets(0).getCellRangeByName("A1:S"+rowSize)
  oArg = Array(lookupIndex, oCellRange, returnIndex, is_sorted) 'array of arguments
  
  For j = 0 To t
	lTick = GetSystemTicks()
		
    If is_sorted = 0 Then
	  OldDoc.sheets(0).getCellByPosition(27,rowIndex).String = oSvc.callFunction("VLOOKUP",oArg)
	End If
	
	If is_sorted = 1 Then
	  OldDoc.sheets(0).getCellByPosition(27,rowIndex).String = oSvc.callFunction("IFERROR",oArg),"Not Found"))
	End If		 	
		 	 
	lTick = (GetSystemTicks() - lTick)
		 
	totalTime = totalTime + lTick
	 
	If lTick > Max Then
	   Max = lTick
	End If
	If lTick < Min Then
	   Min = lTick
	End If
  Next j

  totalTime = totalTime - Max - Min
  
  'write results back to oDoc
  oSheet.getCellByPosition(0, rowIndex).String = rowSize 
  If is_sorted = 0 Then
    oSheet1.getCellByPosition(1, rowIndex).String = totalTime/ 8
  End If
  
  If is_sorted = 1 Then
    oSheet1.getCellByPosition(2, rowIndex).String = totalTime/ 8
  End If
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
  oSheet.getCellByPosition(1, 0).String = "Sorted=False Time (ms)" 
  oSheet.getCellByPosition(2, 0).String = "Sorted=True Time (ms)" 

  rowIndex = 1 'row id where the current result will be written
  
  `iterate over all spreadsheets for is_sorted = 0
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i, 0)
    rowIndex = rowIndex + 1   
  Next i

  `iterate over all spreadsheets for is_sorted = 1
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i, 1)
    rowIndex = rowIndex + 1   
  Next i
End Sub