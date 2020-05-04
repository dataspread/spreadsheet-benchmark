VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

`Takes in reference to a result sheet, row index in the result sheet where
`result will be written, and spreadsheet size (# of spreadsheet rows). 
`Then repeates the experiment for t trials and average memory usage to the results sheet. 
`The average excludes the max and min trial usages for that spreadsheet size.

Private Function GetMemUsage(ByVal rowIndex, ByVal rowSize, ByVal isFormula)
  ' Returns the current Excel.Application
  ' memory usage in MB
  Dim directory As String
  Dim Name As String
  Dim tmpMemUsage As Long
  Dim memUsage As Long
  Dim minUsage As Long
  Dim maxUsage As Long
  
  If isFormula = 0 Then
    'RELATIVE_PATH ---> assign value files directory path here
  End If

  If isFormula = 1 Then
    'RELATIVE_PATH ---> assign formula files directory path here
  End If
  
  'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
    
  Name = FILE_PREFIX & CStr(rowSize) & ".xlsx"
  FILE_PATH = Dir(RELATIVE_PATH & Name)
  Workbooks.Open (FILE_PATH)
    
  minUsage = 10000000
  maxUsage = -1
  t = 10 `10 trials

  For i = 0 To t
    Workbooks.Open (FILE_PATH)    
    Set objSWbemServices = GetObject("winmgmts:")
    
    tmpMemUsage = objSWbemServices.Get( _
      "Win32_Process.Handle='" & _
      GetCurrentProcessId & "'").WorkingSetSize / 1024
      
    memUsage = memUsage + tmpMemUsage
    Set objSWbemServices = Nothing
    Workbooks(FILE_PATH).Close
    
    If tmpMemUsage < minUsage Then
        minUsage = tmpMemUsage
    End If
    If tmpMemUsage > maxUsage Then
        maxUsage = tmpMemUsage
    End If
  Next i
  
  Sheets("Sheet1").Cells(rowIndex, 1).Value = rowSize 
  Sheets("Sheet1").Cells(rowIndex, isFormula + 2).Value = (memUsage - maxUsage - minUsage) / 1024 / 8
  
End Function

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
  Sheets("Sheet1").Cells(1, 2).Value = "Value Size (MB)"
  Sheets("Sheet1").Cells(1, 3).Value = "Formula Size (MB)"

  rowIndex = 2 'row id where the current result will be written
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    GetMemUsage rowIndex, j, 0
    rowIndex = rowIndex + 1     
  Next i

  rowIndex = 2 'row id where the current result will be written
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    GetMemUsage rowIndex, j, 1
    rowIndex = rowIndex + 1     
  Next i
  
  'display results in the result screen
  Application.ScreenUpdating = True
End Sub




