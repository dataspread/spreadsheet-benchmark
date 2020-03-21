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


Function calculateRunTime(ByVal rowIndex, ByVal rowSize)
    
'Initialize Timer Variables (seconds version)
    Dim StartTime As Double
    Dim SecondsElapsed As Double
     
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    
    MaxOne = -1
    MinOne = 1000000
    totalTimeOne = 0
    For k = 1 To 10
    
        timeTaken = 0
        timeOne = 0
        For J = 1 To 5
              
            StartTime = GetTickCount
            result = Application.WorksheetFunction.CountIf(Range("J2:J" & rowSize), "1")
            'Determine how many seconds code took to run
            SecondsElapsed = Round(GetTickCount - StartTime, 2)
            Cells(rowIndex + 50, 28 + J - 1).Value = SecondsElapsed
            If J > 0 Then
                timeTaken = timeTaken + SecondsElapsed
            End If
            
            If J = 1 Then
                timeOne = timeOne + SecondsElapsed
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
        If timeTakenOne > Max Then
          Max = timeTakenOne
        End If
        If timeTakenOne < MinOne Then
          MinOne = timeTakenOne
        End If
        
        
    Next k
    
    
    Cells(rowIndex, 27).Value = rowSize
    Cells(rowIndex, 28).Value = (totalTimeOne - MaxOne - MinOne) / 8
    Cells(rowIndex, 29).Value = (totalTime - Max - Min) / 8
    
End Function


Sub main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'add headers to the current /xlsm file where results will be written
    Sheets("import").Cells(1, 1).Value = "Rows Imported"
    Sheets("import").Cells(1, 2).Value = "Time (s)"

    rowIndex = 2
    For i = 10000 to 500001 Step 10000
        calculateRunTime rowIndex, i
        rowIndex = rowIndex + 1 
    Next i

    Application.ScreenUpdating = True
End Sub





