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
    
    
    Set rg = Range("J2:J" & CStr(rowSize + 1))
    
    For j = 1 To 10
        
        StartTime = Timer
        Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "0")
    
        With cond1
            .Interior.Color = vbGreen
            '.Font.Color = vbWhite
        End With
 
        SecondsElapsed = Round(Timer - StartTime, 2)
        
        timeTaken = SecondsElapsed
        totalTime = totalTime + timeTaken
          
        If timeTaken > Max Then
          Max = timeTaken
        End If
        If timeTaken < Min Then
          Min = timeTaken
        End If
        
        
        'clear any existing conditional formatting
        rg.FormatConditions.Delete
          
          
    Next j
    
    
    totalTime = totalTime - Max - Min
    
    Sheets("import").Activate
   
    Sheets("import").Cells(rowIndex, 1).Value = rowSize
    
    Sheets("import").Cells(rowIndex, 2).Value = totalTime / 8

    
End Sub

Sub main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("import").Cells(1, 1).Value = "Rows Imported"
    Sheets("import").Cells(1, 2).Value = "Time (ms)"
    
    rowIndex = 2
    For i = 10000 to 500001 Step 10000
        calculateRunTime rowIndex, i
        rowIndex = rowIndex + 1 
    Next i
    
    Application.ScreenUpdating = True
End Sub

