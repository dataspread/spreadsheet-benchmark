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

Function calculateRunTimeLinearScale(ByVal rowIndex, ByVal rowSize)
    
'Initialize Timer Variables (seconds version)
    Dim StartTime As Double
    Dim SecondsElapsed As Double
     
    Dim timeArray(10) As Double
    Dim Min(10) As Double
    Dim Max(10) As Double
    For i = 1 To 10
        timeArray(i) = 0
        Min(i) = 1000000
        Max(i) = -1
    Next i

 'Remember time when macro starts
    For k = 1 To 10
        timeTaken = 0
        
        For J = 1 To 1000
            StartTime = GetTickCount
            Cells(rowIndex, 28).Value = Application.WorksheetFunction.CountIf(Range("J2:J" & rowSize), "1")
            'Determine how many seconds code took to run
            SecondsElapsed = Round(GetTickCount - StartTime, 2)
            'Cells(rowIndex + 50, 28 + j - 1).Value = SecondsElapsed
            timeTaken = timeTaken + SecondsElapsed
            
            If J Mod 100 = 0 Then
                If timeTaken > Max(J / 100) Then
                  Max(J / 100) = timeTaken
                End If
                
                If timeTaken < Min(J / 100) Then
                  Min(J / 100) = timeTaken
                End If
                
                timeArray(J / 100) = timeTaken
            
            End If
     
        Next J
    Next k
    
    
    Cells(rowIndex, 27).Value = rowSize
    For i = 1 To 10
        Cells(rowIndex, 38 + i - 1).Value = (timeArray(i)) / 10
    Next i

    
End Function


Sub Test2()
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

    Cells(1, 27).Value = "Rows Read"
    Cells(1, 28).Value = "countif-State"
    Cells(1, 29).Value = "countif-time(ms)"
    For i = 1 To 10
        Cells(1, 38 + i).Value = i & "00 formula"
    Next i

'Calculate Run Time for 150 rows


'last row number of sheet
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
'Foor loop
    'For i = 0 To 1
     '   calculateRunTime i + 2, rowArray(i)
          
    'Next i
    
    J = 2
    For i = 500000 To 500001 Step 10000
        'calculateRunTime j, i
        calculateRunTimeLinearScale J, i
        J = J + 1
    Next i

End Sub


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

    Cells(1, 27).Value = "Rows Read"
    Cells(1, 28).Value = "Single Time (ms)"
    Cells(1, 29).Value = "Multiple Time (ms)"

'Calculate Run Time for 150 rows


'last row number of sheet
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
'Foor loop
    'For i = 0 To 1
     '   calculateRunTime i + 2, rowArray(i)
          
    'Next i
    
    J = 2
    For i = 10000 To 500001 Step 10000
        calculateRunTime J, i
        J = J + 1
    Next i

End Sub





