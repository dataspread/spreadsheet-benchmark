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

Function rangeAccess(ByVal rowIndex, ByVal rowSize)
    
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


Function randomColumnAccess(ByVal rowIndex, ByVal rowSize)
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

Sub main() 
    rowIndex = 2
    For i = 100000 To 500001 Step 200000
        rangeAccess rowIndex, i
        randomColumnAccess rowIndex, i
        rowIndex = rowIndex + 1
    Next i
End Sub
