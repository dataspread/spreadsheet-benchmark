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

Sub calculateRunTimeFalse(ByVal rowIndex, ByVal rowSize, ByVal is_sorted)
    
'Initialize Timer Variables (seconds version)
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim Name As String
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
    
    result = 0
    For j = 1 To 10
          
        StartTime = GetTickCount
        
        If is_sorted = 0 Then
            result = Application.IfError(Application.VLookup(181964, Range("A1:O" & rowSize), 3, False), "Not Found")
        End If
        
        If is_sorted = 1 Then
            result = Application.IfError(Application.VLookup(181964, Range("A1:O" & rowSize), 3, True), "Not Found")
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
          
                 
          
    Next j
    
    totalTime = totalTime - Max - Min
    Sheets("import").Activate
    Sheets("import").Cells(rowIndex, 1).Value = rowSize
    If is_sorted = 0 Then
	Sheets("import").Cells(rowIndex, 2).Value = totalTime / 8
    End If

    If is_sorted = 1 Then
	Sheets("import").Cells(rowIndex, 3).Value = totalTime / 8
    End If
    
    
End Sub

Sub main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("import").Cells(1, 1).Value = "Rows Imported"
    Sheets("import").Cells(1, 2).Value = "False Time (ms)"
    Sheets("import").Cells(1, 3).Value = "True Time (ms)"
    
    rowIndex = 2
    For i = 10000 to 500001 Step 10000
        calculateRunTime rowIndex, i, 0
        rowIndex = rowIndex + 1 
    Next i
    
    rowIndex = 2
    For i = 10000 to 500001 Step 10000
        calculateRunTime rowIndex, i, 1
        rowIndex = rowIndex + 1 
    Next i
    
    Application.ScreenUpdating = True
End Sub

