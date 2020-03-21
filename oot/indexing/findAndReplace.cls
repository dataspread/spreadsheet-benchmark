VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub calculateRunTime(ByVal rowIndex, ByVal rowSize)
    
'Initialize Timer Variables
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim Name As String
    
    Dim rg As Range
    Dim cond1 As FormatCondition
    
    Dim sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    
    'RELATIVE_PATH ---> assign directory path here
    'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
    
    Name = FILE_PREFIX & CStr(rowSize) & ".xlsx"
    FILE_PATH = Dir(RELATIVE_PATH & Name)
    sheetName = "Sheet1"
    
    Workbooks.Open (FILE_PATH)
        
    Workbooks(FILE_PATH).Worksheets(sheetName).Range("A1:O" & CStr(rowSize + 1)).Copy
    
    'Activate the destination worksheet
    Sheets("Sheet1").Activate
    'Select the target range
    Range("A1").Select
    'Paste in the target destination
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    Workbooks(FILE_PATH).Close
    
    totalTimeS = 0
    totalTimeR = 0
        
    For j = 1 To 10
    
        Search = "yawwn"
        Replacement = "wohoo"
          
        StartTime = Timer
        Worksheets("Sheet1").Range("A1:O" & CStr(rowSize + 1)).Replace What:=Search, Replacement:=Replacement, _
                SearchOrder:=xlByRows, MatchCase:=False
        SecondsElapsed = Round(Timer - StartTime, 2)
        
        totalTimeS = totalTimeS + SecondsElapsed
        
        
        If SecondsElapsed > MaxS Then
            MaxS = SecondsElapsed
        End If
        If SecondsElapsed < MinS Then
          MinS = SecondsElapsed
        End If
        
        Search = "wohoo"
        Replacement = "yawwn"
        
        StartTime = Timer
        
        Worksheets("Sheet1").Range("A1:O" & CStr(rowSize + 1)).Replace What:=Search, Replacement:=Replacement, _
                SearchOrder:=xlByRows, MatchCase:=False
        SecondsElapsed = Round(Timer - StartTime, 2)
        
        totalTimeR = totalTimeR + SecondsElapsed
        
        If SecondsElapsed > MaxR Then
            MaxR = SecondsElapsed
        End If
        If SecondsElapsed < MinR Then
          MinR = SecondsElapsed
        End If
                    
          
    Next j
    
    totalTimeS = totalTimeS - MaxS - MinS
    totalTimeR = totalTimeR - MaxR - MinR
    Sheets("import").Activate
   
    Sheets("import").Cells(rowIndex, 1).Value = rowSize
    Sheets("import").Cells(rowIndex, 2).Value = totalTimeS / 8
    Sheets("import").Cells(rowIndex, 3).Value = totalTimeR / 8
End Sub

Sub calculateRunTimeAbsentValueSearch(ByVal rowIndex, ByVal rowSize)
    
'Initialize Timer Variables
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim Name As String
    
    Dim rg As Range
    Dim cond1 As FormatCondition
    
    Dim sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
    
    
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    
    'RELATIVE_PATH ---> assign directory path here
    'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
    
    Name = FILE_PREFIX & CStr(rowSize) & ".xlsx"
    FILE_PATH = Dir(RELATIVE_PATH & Name)
    sheetName = "Sheet1"
    
    Workbooks.Open (FILE_PATH)
        
    Workbooks(FILE_PATH).Worksheets(sheetName).Range("A1:O" & CStr(rowSize + 1)).Copy
    
    'Activate the destination worksheet
    Sheets("Sheet1").Activate
    'Select the target range
    Range("A1").Select
    'Paste in the target destination
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    Workbooks(FILE_PATH).Close
    
    totalTimeS = 0
    totalTimeR = 0
        
    For j = 1 To 10
    
        Search = "eeeeehaaaaaaaaaaa"
        Replacement = "ayyyyyhayyyyyy"
          
        StartTime = Timer
        Worksheets("Sheet1").Range("A1:O" & CStr(rowSize + 1)).Replace What:=Search, Replacement:=Replacement, _
                SearchOrder:=xlByRows, MatchCase:=False
        SecondsElapsed = Round(Timer - StartTime, 2)
        
        totalTimeS = totalTimeS + SecondsElapsed
        
        
        If SecondsElapsed > MaxS Then
            MaxS = SecondsElapsed
        End If
        If SecondsElapsed < MinS Then
          MinS = SecondsElapsed
        End If
    Next j
    
    totalTimeS = totalTimeS - MaxS - MinS
    Sheets("import").Activate
    Sheets("import").Cells(rowIndex, 4).Value = totalTimeS / 8
End Sub

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

    rowIndex = 2
    For i = 10000 to 500001 Step 10000
        calculateRunTimeAbsentValueSearch rowIndex, i
        rowIndex = rowIndex + 1 
    Next i

    Application.ScreenUpdating = True
End Sub




