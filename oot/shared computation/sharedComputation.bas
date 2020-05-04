REM  *****  BASIC  *****

'put formulae
Function putFormula(numOfFormula As Long, oSheet1 As Object, strategy As Long)
  For i = 0 to 9
	  oSheet1.getCellRangeByName("U1:U" & numOfFormula).ClearContents(2 ^ i)
  Next i
  For i = 1 To numOfFormula
    
  Next i
  For i = 1 To numOfFormula
    If strategey = 0 Then 'repeated computation
      oSheet1.getCellRangeByName("U"&i).Formula = "=SUM(J$2:J$" & (i+1) & ")"
    End If
    
    If strategey = 1 Then 'resuable computation
      If i=0 Then
        oSheet1.getCellRangeByName("U"&1).Formula = "=SUM(J$2:J$" & (2) & ")"
      Else
        oSheet1.getCellRangeByName("U"&i).Formula = "=SUM(U" & (i-1) & "; J" & (i+1) & ")"
      End If 
    End If
  Next i
End Function

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTime(rowIndex As Long, numOfFormula As Long, strategy As Long)
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

  putFormula(numOfFormula, OldDoc.sheets(0), strategy)

  For j = 0 To t
    lTick = GetSystemTicks()
    ThisComponent.calculateAll()
    lTick = (GetSystemTicks() - lTick)
    
    totalTime = totalTime + lTick
		     
	If lTick > Max Then
	  Max = lTick
	End If
	
    If lTick < Min Then
	    Min = lTick
	  End If
  Next j

  OldDoc.dispose 'close document  
  totalTime = totalTime - Max - Min 'remove outliers
  
  'write results back to oDoc
  oSheet.getCellByPosition(27,rowIndex).String = rowSize
  oSheet.getCellByPosition(26, rowIndex).String = totalTime/8
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
  ThisComponent.isAutomaticCalculationEnabled = False
  minRows = 10000 `min row size
  maxRows = 100000 `max row size
  stepSize = 10000 `increment row sizes by 10k
  
  'add headers to the current file where results will be written
  oSheet.getCellByPosition(0,0).String = "Import Size" 
  oSheet.getCellByPosition(1, 0).String = "Time (ms)" 

  rowIndex = 1 'row id where the current result will be written
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i, 0)
    rowIndex = rowIndex + 1   
  Next i

  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i, 1)
    rowIndex = rowIndex + 1   
  Next i
End Sub
