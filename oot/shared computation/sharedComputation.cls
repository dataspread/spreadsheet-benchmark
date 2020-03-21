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

Function calculateRepeatedComputationTime(ByVal rowid, ByVal formulaCount)
    
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
        If Val(Application.Version) >= 12 Then
            Range("P1:P" & CStr(formulaCount)).CalculateRowMajorOrder
        Else
            Range("P1:P" & CStr(formulaCount)).Calculate
        End If
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
    Cells(rowid, 28).Value = totalTime / 8
    
End Function

Function sharedReCalculateRunTime(ByVal rowid, ByVal formulaCount)
    
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
        
        
        If Val(Application.Version) >= 12 Then
            Range("P1:P" & CStr(formulaCount)).CalculateRowMajorOrder
        Else
            Range("P1:P" & CStr(formulaCount)).Calculate
        End If
    'Determine how many seconds code took to run
        SecondsElapsed = GetTickCount - StartTime
        
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
    Cells(rowid, 29).Value = totalTime / 8
    
End Function

Function putFormula(ByVal formulaCount)
    Range("P1:P10000").Clear
    For i = 1 To formulaCount
        Worksheets("Sheet1").Range("P" & CStr(i)).Formula = "=SUM($J$2:$J$" & CStr(i + 1) & ")"
    Next i
End Function

Function putFormulaShared(ByVal formulaCount)
    Range("P1:P10000").Clear
    Worksheets("Sheet1").Range("P1").Formula = "=SUM($J$2:$J$2)"
    For i = 2 To formulaCount
        Worksheets("Sheet1").Range("P" & CStr(i)).Formula = "=SUM(P" & CStr(i - 1) & ",J" & CStr(i + 1) & ")"  
    Next i
End Function

Function repeatedComputationTime()
    Application.Calculation = xlManual
    rowIndex = 2
    
    For J = 10000 To 100001 Step 10000
        putFormula J
        calculateRepeatedComputationTime i, J
        rowIndex = rowIndex + 1       
    Next J
End Function


Function reusableComputationTime()
    Application.Calculation = xlManual
    rowIndex = 2
    
    For J = 10000 To 100001 Step 10000        
        putFormulaShared J
        calculateReusableComputationTime i, J
        rowIndex = rowIndex + 1
        
    Next J
End Function

Sub main()
    repeatedComputationTime
    reusableComputationTime
End Sub


