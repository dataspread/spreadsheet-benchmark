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

Function fullCalculateRunTime(ByVal rowid)
    
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
        Application.CalculateFull
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
    Cells(rowid, 29).Value = totalTime / 8
    
End Function

Function putFormula(ByVal rows)
    Dim c As Range

    Set c = Range("J2:J" & CStr(rows))

      
    Range("P1:P10000").Clear
    For i = 1 To 100

        Worksheets("Sheet1").Range("P" & CStr(i)).Formula = "=COUNTIF($J$2:$J$" & CStr(rows) & ",""1"")"
    
    Next i

End Function
Sub Test1()
'PURPOSE: Determine how many seconds it took for code to completely run
    
'Initialize array of row sizes
    
    Dim rowArray As Variant
    rowArray = Array(150, 6000)
'Declare row iterator i

    Dim i As Long
    Dim r As Long
'Declare end row number counter

    Dim lastRow As Long
'set headers at the first row. 17,18, and 19 are the coulmn numbers

    Application.Calculation = xlManual

    Cells(1, 27).Value = "Rows Read"
    Cells(1, 28).Value = "Recalc Time(ms)"

'Calculate Run Time for 150 rows
    i = 2
    putFormula 1
    
    For J = 10000 To 500001 Step 10000
        Cells(i, 27).Value = J
        putFormula J
        reCalculateRunTime i
        'fullCalculateRunTime i
        i = i + 1
        'fullCalculateRunTime
    Next J
'last row number of sheet
    lastRow = Range("A" & rows.Count).End(xlUp).Row
'Foor loop
    'For i = 0 To 1
     '   calculateRunTime i + 2, rowArray(i)
          
    'Next i
    
    
    

End Sub



