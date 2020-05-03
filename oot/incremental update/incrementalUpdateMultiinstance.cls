VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private mlngStart As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'recalculate all the formula instances for the given spreadsheet
Function reCalculateRunTime(ByVal rowid)
'Initialize Timer Variables (milliseconds version, more error prone I believe)
  Dim StartTime As Double
  Dim SecondsElapsed As Double
  
 'Remember time when macro starts
  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials
  
  MaxU = -1
  MinU = 1000000
  totalTimeU = 0
  
  Cells(2, 10).Value = 1 'initialize cell value to 1
  For J = 0 To t
      
    StartTime = GetTickCount    
    Cells(2, 10).Value = 0 'update value
    Application.Calculate 'rcalculate
    SecondsElapsed = Round(GetTickCount - StartTime, 2)
    
    timeTaken = SecondsElapsed
    totalTime = totalTime + timeTaken
      
    If timeTaken > Max Then
      Max = timeTaken
    End If
    If timeTaken < Min Then
      Min = timeTaken
    End If   
    Cells(2, 10).Value = 1
      
  Next J
  
  totalTime = totalTime - Max - Min
  'write results back to result sheet
  Sheets("Sheet1").Cells(rowIndex, 27).Value = rowSize   
  Sheets("Sheet1").Cells(rowIndex, 28).Value = totalTime / 8 
  
End Function

'set the formulae
Function putFormula(ByVal counts, ByVal rows)
  Range("P1:P" & CStr(counts)).Clear
  For i = 1 To counts+1
    Worksheets("Sheet1").Range("P" & CStr(i)).Formula = "=COUNTIF($J$2:$J$" & CStr(rows) & ",""1"")"
  Next i
End Function

'Runs experiments on all instances for the given spreadsheet
'specified by [minCount, maxCount] with stepSize increments.
'This is the main function to be called for running the experiment.

Sub main()
  Dim i As Long
  Dim minCount as Long
  Dim maxCount as Long
  Dim stepSize as Long
  Dim rowIndex as Long

  minCount = 100 `min row size
  maxCount = 1000 `max row size
  stepSize = 100 `increment row sizes by 10k

  'prevent screen updating when the macro is runnign
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  
  Application.Calculation = xlManual 'set manual calculation mode

  'add headers to the Sheet1 of result file (the .xlsm file)
  Sheets("Sheet1").Cells(1, 1).Value = "Rows Imported"
  Sheets("Sheet1").Cells(1, 2).Value = "Time (s)"

  rowIndex = 2 'row id where the current result will be written
  
  `iterate over all instances
  totalRows = 500000
  For J = minCount To maxCount+1 Step stepSize
    putFormula J, totalRows
    reCalculateRunTime rowIndex
    rowIndex = rowIndex + 1
  Next J
End Sub



