REM  *****  BASIC  *****

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
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

  MaxOne = -1
  MinOne = 1000000
  totalTimeOne = 0
  t = 10 `10 trials
  numInstances = 5 'no of redundant instances

  'RELATIVE_PATH ---> assign directory path here
  'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
  FILE_PATH = RELATIVE_PATH & "/" & FILE_PREFIX & "/" & (rowSize) & ".ods"
  url = ConvertToURL(FILE_PATH)
  OldDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open document

  oSvc = createUnoService( "com.sun.star.sheet.FunctionAccess")
  condition = "1"
  oCellRange = OldDoc.sheets(0).getCellRangeByName("J2:J"&rowSize)
  oArg = Array(oCellRange, condition) 'array of arguments

  For k = 1 To t
    timeTaken = 0
    timeOne = 0
    For J = 1 To numInstances
      lTick = GetSystemTicks()
      OldDoc.sheets(0).getCellByPosition(27+J,rowIndex).String = oSvc.callFunction( "COUNTIFS", oArg)
      lTick = (GetSystemTicks() - lTick)
      If J > 0 Then
        timeTaken = timeTaken + lTick
      End If
      
      If J = 1 Then
        timeOne = timeOne + lTick
      End If
    Next J
    
    totalTime = totalTime + timeTaken
    If timeTaken > Max Then
      Max = timeTaken
    End If
    If timeTaken < Min Then
      Min = timeTaken
    End If
    
    totalTimeOne = totalTimeOne + timeOne
    If timeOne > MaxOne Then
      MaxOne = timeOne
    End If
    If timeOne < MinOne Then
      MinOne = timeOne
    End If
  Next k
  
  OldDoc.dispose 'close document  
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
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i)
    rowIndex = rowIndex + 1   
  Next i
End Sub
