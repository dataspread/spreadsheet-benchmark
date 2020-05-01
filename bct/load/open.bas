REM  *****  BASIC  *****

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
  Dim openedDoc As Object 
  Dim totalTime As Long
  Dim Max As Long
  Dim Min As Long
  Dim t As Long
  Dim Prop(0) as New com.sun.star.beans.PropertyValue

  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials

  'RELATIVE_PATH ---> assign directory path here
  'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
  FILE_PATH = RELATIVE_PATH & "/" & FILE_PREFIX & "/" & (rowSize) & ".ods"
  url = ConvertToURL(FILE_PATH)

  For j = 0 To t
	lTick = GetSystemTicks()
	openedDoc = stardesktop.LoadComponentFromURL(url, "_blank",0, Prop()) 'open document
	lTick = (GetSystemTicks() - lTick)
	openedDoc.dispose 'close document
	oSheet.getCellByPosition(2+j,rowIndex).String = lTick
	totalTime = totalTime + lTick	 
	If lTick > Max Then
	   Max = lTick
	End If
	
	If lTick < Min Then
	   Min = lTick
	End If
  Next j

  totalTime = totalTime - Max - Min 'remove outliers
  
  'write results back to oDoc
  oSheet.getCellByPosition(0,rowIndex).String = rowSize 
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
  For i = minRows to maxRows+1 Step 10000
    calculateRunTime(oDoc,oSheet,rowIndex,i)
    rowIndex = rowIndex + 1   
  Next i

End Sub
