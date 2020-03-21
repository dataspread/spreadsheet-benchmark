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
Function reCalculateRunTime(ByVal rowid)
    
'Initialize Timer Variables (milliseconds version, more error prone I believe)
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    
    MaxU = -1
    MinU = 1000000
    totalTimeU = 0
        
    For J = 1 To 10
          
        StartTime = GetTickCount
        'Cells(rowIndex, 28).Value = Application.WorksheetFunction.CountIf(Range("J2:J" & rowSize), "1")
        
        Cells(2, 10).Value = 0
        Application.Calculate
    'Determine how many seconds code took to run
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
    Cells(rowid, 28).Value = totalTime / 8
    
End Function

Function putFormula(ByVal formulaCount)
    Range("P1:P10000").Clear
    For i = 1 To formulaCount
        Worksheets("Sheet1").Range("P" & CStr(i)).Formula = "=COUNTIF($J$2:$J$500001,""1"")"
    Next i

End Function

Sub main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'add headers to the current /xlsm file where results will be written
    Sheets("import").Cells(1, 1).Value = "Rows Imported"
    Sheets("import").Cells(1, 2).Value = "Time (s)"

    rowIndex = 2
    For J = 10000 To 500001 Step 10000
        putFormula J
        reCalculateRunTime rowIndex
        rowIndex = rowIndex + 1
    Next J
End Sub



