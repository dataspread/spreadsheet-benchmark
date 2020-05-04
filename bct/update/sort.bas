REM  *****  BASIC  *****

'set sort property
Sub SortCells( curSheet as Object, rowsize As Long, sortAscending As Boolean) 
   ' An array of sort fields determines the columns that are sorted.
  Dim sortFields(0) As New com.sun.star.util.SortField
   
   ' The sort descriptor is an array of properties.
   ' The primary property contains the sort fields.
   Dim sortDesc(1) As New com.sun.star.beans.PropertyValue

   ' The columns are numbered starting with 0, so column A is 0, column B is 1, etc...
   sortFields(0).Field = 0
   sortFields(0).SortAscending = sortAscending
   
   ' Setup the sort descriptor.
   sortDesc(0).Name = "SortFields"
   sortDesc(0).Value = sortFields()
   sortDesc(1).Name = "HasHeader" ' indicate data contains header row
   sortDesc(1).Value = false
   
   ' Sort the range.
   curSheet.getCellRangeByPosition(0,1,15,rowsize).Sort(sortDesc())
End Sub 

Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
  Dim OldDoc As Object
  Dim oRange As Object 
  Dim totalTime As Long
  Dim Max As Long
  Dim Min As Long
  Dim curSheet as object, oCtrl as object
  Dim sortAscending As Boolean 'sort dataset in ascending order

  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials
   
  'RELATIVE_PATH ---> assign directory path here
  'FILE_PREFIX ---> assuming all the files in the directory have a common prefix followed by its number of rows
  FILE_PATH = RELATIVE_PATH & "/" & FILE_PREFIX & "/" & (rowSize) & ".ods"
  url = ConvertToURL(FILE_PATH)
  
  For j = 0 To t    
    OldDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open document
    curSheet = OldDoc.sheets(0)
    sortAscending = true
    lTick = GetSystemTicks()   
	  SortCells (curSheet, rowSize, sortAscending)
    lTick = (GetSystemTicks() - lTick)
 
    totalTime = totalTime + lTick
     
    If lTick > Max Then
       Max = lTick
    End If
     
    If lTick < Min Then
       Min = lTick
    End If
    sortAscending = false    
	  SortCells (curSheet, rowSize, sortAscending) 'unsort the data
	  OldDoc.dispose 'close document
  Next j
  
  totalTime = totalTime - Max - Min
  
  'write results back to oDoc
  oSheet.getCellByPosition(0,rowIndex).String = rowSize
  oSheet.getCellByPosition(1, rowIndex).String = totalTime/8
End Sub

'Runs experiments on all spreadsheets specified by  [minRows, maxRows] with stepSize increments.
'This is the main function to be called for running the experiment.

Sub main
  Dim oDoc As Object
  Dim oSheet As Object
  Dim minRows as Long
  Dim maxRows as Long
  Dim stepSize as Long
  Dim rowIndex as Long
  
  oDoc = ThisComponent ' the file where the results are written
  oSheet = oDoc.Sheets(0) 'get first sheet
  minRows = 10000 `min row size
  maxRows = 50000 `max row size
  stepSize = 10000 `increment row sizes by 10k
  
  'add headers to the current file where results will be written
  oSheet.getCellByPosition(0,0).String = "Import Size" 
  oSheet.getCellByPosition(1, 0).String = "Time (ms)" 

  rowIndex = 1 'row id where the current result will be written
  
  `iterate over all spreadsheets
  For i = minRows to maxRows+1 Step stepSize
    calculateRunTime(oDoc,oSheet,rowIndex,i)
    rowIndex = rowIndex + 1   
  Next i
End Sub