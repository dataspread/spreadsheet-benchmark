VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub calculateRunTime(ByVal rowIndex, ByVal rowCount)
    
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
    
    Dim pf As String
    Dim pf_Name As String
    
    rowSize = rowCount * 1000
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    
    directory = "C:\Users\rzhang74\Desktop\Research2018Fall\Spring2020\open\"
    Name = "airbnb_" & CStr(rowCount) & "k.xlsx"
    sheetName = "airbnb_" & CStr(rowCount) & "k"
    Filename = Dir(directory & Name)
    
    Sheets("Sheet1").Delete
    Dim sheetNew As Worksheet
    Set sheetNew = Sheets.Add
    sheetNew.Name = "Sheet1"

    
    Workbooks.Open (directory & Filename)
        
    Workbooks(Filename).Worksheets(sheetName).Range("A1:O" & CStr(rowSize + 1)).Copy
        
    'Activate the destination worksheet
    Sheets("Sheet1").Activate
    'Select the target range
    Range("A1").Select
    'Paste in the target destination
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    Workbooks(Filename).Close
    
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
Sub calculateRunTimeNE(ByVal rowIndex, ByVal rowCount)
    
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
    
    Dim pf As String
    Dim pf_Name As String
    
    rowSize = rowCount * 1000
 'Remember time when macro starts
    Max = -1
    Min = 1000000
    totalTime = 0
    
    directory = "C:\Users\rzhang74\Desktop\Research2018Fall\Spring2020\open\"
    Name = "airbnb_" & CStr(rowCount) & "k.xlsx"
    sheetName = "airbnb_" & CStr(rowCount) & "k"
    Filename = Dir(directory & Name)
    
    Sheets("Sheet1").Delete
    Dim sheetNew As Worksheet
    Set sheetNew = Sheets.Add
    sheetNew.Name = "Sheet1"

    
    Workbooks.Open (directory & Filename)
        
    Workbooks(Filename).Worksheets(sheetName).Range("A1:O" & CStr(rowSize + 1)).Copy
        
    'Activate the destination worksheet
    Sheets("Sheet1").Activate
    'Select the target range
    Range("A1").Select
    'Paste in the target destination
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    Workbooks(Filename).Close
    
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

Sub Test1()
'PURPOSE: Determine how many seconds it took for code to completely run
    
'Initialize array of row sizes
    
   

    Dim i As Long

'set headers at the first row. 17,18, and 19 are the coulmn numbers
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("import").Cells(1, 1).Value = "Rows Imported"
    Sheets("import").Cells(1, 2).Value = "Time (s)"
    Sheets("import").Cells(1, 3).Value = "Reverse Time (s)"
    Sheets("import").Cells(1, 4).Value = "No exist Time (s)"
    
    i = 2
    For j = 10 To 501 Step 10
        'calculateRunTime i, j
        calculateRunTimeNE i, j
        i = i + 1
          
    Next j
    
    Application.ScreenUpdating = True
End Sub




