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

Function sequentialAccess(ByVal rowIndex, ByVal rowSize)
    
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

    Cells(rowIndex, 31).Value = rowSize
    Cells(rowIndex, 32).Value = (totalTime - Max - Min) / 8
    
End Function

Function randomAccess(ByVal rowIndex, ByVal rowSize)
    
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
    Cells(rowIndex, 33).Value = (totalTime - Max - Min) / 8
    
End Function

Sub main() 
    rowIndex = 2
    For i = 100000 To 500001 Step 200000
        sequentialAccess rowIndex, i
        randomAccess rowIndex, i
        rowIndex = rowIndex + 1
    Next i
End Sub



