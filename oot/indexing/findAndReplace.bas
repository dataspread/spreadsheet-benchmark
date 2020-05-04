REM  *****  BASIC  *****

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, 
                     rowSize As Long, search as String, replacement as String, columnIndex as Long)
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

  oCellRange = OldDoc.sheets(0).getCellRangeByName("A1:P"+rowSize)
  replace = OldDoc.sheets(0).createReplaceDescriptor rem OldDoc.createReplaceDescriptor in case of Writer
  rem replace.SearchRegularExpression = True
  
  For j = 0 To t
    lTick = GetSystemTicks()
    replace.SearchString = search
	replace.ReplaceString = replacement
	oCellRange.replaceAll(replace)
    lTick = (GetSystemTicks() - lTick)
	       
	totalTime = totalTime + lTick
	         
	If lTick > Max Then
           Max = lTick
	End If
	If lTick < Min Then
           Min = lTick
	End If

	replace.SearchString = replacement
	replace.ReplaceString = search
	oCellRange.replaceAll(replace) 'reset spreadsheet
  Next j
    
  OldDoc.dispose 'close document  
  totalTime = totalTime - Max - Min 'remove outliers
  
  'write results back to oDoc
  oSheet.getCellByPosition(26,rowIndex).String = rowSize
  oSheet.getCellByPosition(26+columnIndex, rowIndex).String = totalTime/8
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
  Search = "yawwn" ' a value present in the spreadsheet
  Replacement = "wohoo" ' the value that replaces the searched value
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i,Search,Replacement,2)
    rowIndex = rowIndex + 1   
  Next i

  rowIndex = 1 'row id where the current result will be written
  Search = "eeeeehaaaaaaaaaaa" ' a value present in the spreadsheet
  Replacement = "ayyyyyhayyyyyy" ' the value that replaces the searched value
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i,Search,Replacement,3)
    rowIndex = rowIndex + 1   
  Next i
End Sub
