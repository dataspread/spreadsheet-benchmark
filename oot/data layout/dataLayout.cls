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
Function calculateRunTimeRowRange(ByVal rowIndex, ByVal rowSize)
    
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
        result = Application.WorksheetFunction.Count(Range("A2:O" & rowSize))
    
        
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
          
                 
          
    Next J
    
    totalTime = totalTime - Max - Min
    Cells(rowIndex, 37).Value = rowSize
    Cells(rowIndex, 38).Value = totalTime / 8
    
End Function


Function calculateRunTimeRowRandCol(ByVal rowIndex, ByVal rowSize)
    Dim InArray(15) As Variant
    lower = 0
    upper = 14
    For i = lower To upper
        InArray(i) = i + 1
    Next i
    
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long
   
    Randomize
    For N = lower To upper
        J = CLng(((upper - N) * Rnd) + N)
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
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
        Sum = 0
        For N = lower To upper
            colIndex = InArray(N)
            colLetter = Split(Cells(1, colIndex).Address, "$")(1)
            Sum = Application.WorksheetFunction.Count(Range(colLetter & "2:" & colLetter & rowSize))
        Next N
        
        
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
          
                 
          
    Next J
    
    totalTime = totalTime - Max - Min
    Cells(rowIndex, 39).Value = totalTime / 8
    
End Function


Function calculateRunTimeColLayoutSeq(ByVal rowIndex, ByVal rowSize)
    
'Initialize Timer Variables (seconds version)
    Dim StartTime As Double
    Dim SecondsElapsed As Double
     
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    timeTaken = 0
    i = 0
    
    'Debug.Print rowSize
    ReDim InArray(rowSize) As Variant
    lower = 0
    upper = rowSize - 1
    For i = lower To upper

        InArray(i) = i + 2
    Next i
    'Debug.Print lower, upper
    For k = 1 To 10
        StartTime = GetTickCount
        For J = lower To upper
            i = i + Cells(InArray(J), 10).Value '+ Cells(InArray(J), 11).Value + Cells(InArray(J), 12).Value
        Next J
        SecondsElapsed = Round(GetTickCount - StartTime, 2)
        timeTaken = SecondsElapsed
        totalTime = totalTime + timeTaken
        If timeTaken > Max Then
          Max = timeTaken
        End If
        If timeTaken < Min Then
          Min = timeTaken
        End If
    Next k
    
    'Debug.Print i
    Cells(rowIndex, 31).Value = rowSize
    Cells(rowIndex, 32).Value = (totalTime - Max - Min) / 8
    
End Function

Function calculateRunTimeColLayoutRand(ByVal rowIndex, ByVal rowSize)
    
'Initialize Timer Variables (seconds version)
    Dim StartTime As Double
    Dim SecondsElapsed As Double
     
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    timeTaken = 0
    i = 0
    
    ReDim InArray(rowSize) As Variant
    lower = 0
    upper = rowSize - 1
    For i = lower To upper
        InArray(i) = i + 2
    Next i
    
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long
   
    Randomize
    For N = lower To upper
        J = CLng(((upper - N) * Rnd) + N)
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
    
    i = 0
    For k = 1 To 10
        StartTime = GetTickCount
        For J = lower To upper
            i = i + Cells(InArray(J), 10).Value ' + Cells(InArray(J), 11).Value + Cells(InArray(J), 12).Value
        Next J
        SecondsElapsed = Round(GetTickCount - StartTime, 2)
        timeTaken = SecondsElapsed
        totalTime = totalTime + timeTaken
        If timeTaken > Max Then
          Max = timeTaken
        End If
        If timeTaken < Min Then
          Min = timeTaken
        End If
    Next k
    
    Debug.Print i
    Cells(rowIndex, 33).Value = (totalTime - Max - Min) / 8
    
End Function

Sub Test1()
 'PURPOSE: Determine how many seconds it took for code to completely run
    
    Dim i As Long
    Dim r As Long
'Declare end row number counter

    Dim lastRow As Long
'set headers at the first row. 17,18, and 19 are the coulmn numbers

    Cells(1, 31).Value = "Rows Read"
    Cells(1, 32).Value = "Col Layout Seq(ms)"
    Cells(1, 33).Value = "Col Layout rand (ms)"
    Cells(1, 37).Value = "Rows Read"
    Cells(1, 38).Value = "Row Layout Range(ms)"
    Cells(1, 39).Value = "Row Layout rand (ms)"

'Calculate Run Time for 150 rows


'last row number of sheet
'lastRow = Range("A" & Rows.Count).End(xlUp).Row
'Foor loop
    'For i = 0 To 1
     '   calculateRunTime i + 2, rowArray(i)
          
    'Next i
    
    J = 2
    For i = 100000 To 500001 Step 200000
        'calculateRunTimeColLayoutSeq J, i
        J = J + 1
    Next i
    
    J = 2
    For i = 100000 To 500001 Step 200000
        calculateRunTimeColLayoutRand J, i
        J = J + 1
    Next i
    
    
    J = 2
    For i = 100000 To 500001 Step 200000
        'calculateRunTimeRowRange J, i
        J = J + 1
    Next i
    
    
    J = 2
    For i = 100000 To 500001 Step 200000
        'calculateRunTimeRowRandCol J, i
        J = J + 1
    Next i
    
    
    
    
End Sub



