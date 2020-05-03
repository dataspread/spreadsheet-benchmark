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

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average time to the results sheet. 
`The average excludes the max and min trial times for that spreadsheet size.

Sub calculateRunTimeFalse(ByVal rowIndex, ByVal rowSize, ByVal is_sorted)
  
'Initialize Timer Variables (seconds version)
  Dim StartTime As Double
  Dim SecondsElapsed As Double
  Dim Name As String
 'Remember time when macro starts
  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials
  
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
  For j = 0 To t
      
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
  
  totalTime = totalTime - Max - Min 'remove outliers
  
  'write results back to result sheet
  Sheets("Sheet1").Activate
  Sheets("Sheet1").Cells(rowIndex, 1).Value = rowSize
  
  If is_sorted = 0 Then
	Sheets("Sheet1").Cells(rowIndex, 2).Value = totalTime / 8
  End If

  If is_sorted = 1 Then
	Sheets("Sheet1").Cells(rowIndex, 3).Value = totalTime / 8
  End If
End Sub

'Runs experiments on all spreadsheets specified by  [minRows, maxRows] with stepSize increments.
'This is the main function to be called for running the experiment.

Sub main() 
  Dim i As Long
  Dim minRows as Long
  Dim maxRows as Long
  Dim stepSize as Long
  Dim rowIndex as Long

  minRows = 10000 `min row size
  maxRows = 50000 `max row size
  stepSize = 10000 `increment row sizes by 10k

  'prevent screen updating when the macro is runnign
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

  'add headers to the Sheet1 of result file (the .xlsm file)
  Sheets("Sheet1").Cells(1, 1).Value = "Rows Sheet1ed"
  Sheets("Sheet1").Cells(1, 2).Value = "Sorted=False Time (s)"
  Sheets("Sheet1").Cells(1, 3).Value = "Sorted=True Time (s)"

  rowIndex = 2 'row id where the current result will be written
  
  `iterate over all spreadsheets for is_sorted = 0
  For i = minRows to maxRows+1 Step stepSize
  calculateRunTime rowIndex, i
  rowIndex = rowIndex + 1   
  Next i
  
  rowIndex = 2 'row id where the current result will be written
  
  `iterate over all spreadsheets for is_sorted = 1
  For i = minRows to maxRows+1 Step stepSize
  calculateRunTime rowIndex, i, 0, 1
  rowIndex = rowIndex + 1   
  Next i
  
  'display results in the result screen
  Application.ScreenUpdating = True
End Sub

