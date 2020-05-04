REM  *****  BASIC  *****

'range access experiment
Sub rangeAccess(rowIndex As Long, rowSize As Long)
  Dim oDoc As Object
  Dim oSheet1 As Object
  Dim oCell As Object
  Dim oCellRange As Object
  Dim oActiveRange As Object
  Dim oSvc as variant
  Dim oArg as variant
  Dim lTick As Long
   
  Dim totalTime As Long
  Dim Max As Long
  Dim Min As Long

  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials

  oSvc = createUnoService( "com.sun.star.sheet.FunctionAccess")
  oDoc = ThisComponent
  oSheet1 = oDoc.Sheets(0)

  Dim sum As Long
  Dim MyArray(rowSize-1) As Long
  lower = 1
  upper = rowSize-1
   
  For j = 0 To t  	
    Max = -1
    Min = 1000000
    sum = 0
    lTick = GetSystemTicks()
    oCellRange = oSheet1.getCellRangeByName("A1:R"+upper)
  	oArg = Array(oCellRange)
    sum =  oSvc.callFunction( "COUNT", oArg)
    
    lTick = (GetSystemTicks() - lTick)
       
    totalTime = totalTime + lTick
    If lTick > Max Then
     Max = lTick
    End If
    If lTick < Min Then
     Min = lTick
    End If
  Next j
   
  totalTime = totalTime - Max - Min
  oSheet1.getCellByPosition(31, rowIndex).String = totalTime/8
End Sub

'random colum access experiment
Sub randomColumnAccess(rowIndex As Long, rowSize As Long)
  Dim oDoc As Object
  Dim oSheet1 As Object
  Dim oCell As Object
  Dim oCellRange As Object
  Dim oActiveRange As Object
  Dim oSvc as variant
  Dim oArg as variant
  Dim lTick As Long
   
  Dim totalTime As Long
  Dim Max As Long
  Dim Min As Long

  Max = -1
  Min = 1000000
  totalTime = 0
  t = 10 `10 trials
	
  oSvc = createUnoService( "com.sun.star.sheet.FunctionAccess")
  oDoc = ThisComponent
  oSheet1 = oDoc.Sheets(0)

  Dim sum As Long
  Dim MyArray(15) As Long
  Dim colArray(15) As String
  lower = 0
  upper = 14
  For i = lower To upper
    MyArray(i) = i
  Next i
  
  Dim Temp as Long
  For i = lower To upper
    j = CLng((upper - i) * rnd + i)
    If i <> j Then
      Temp = MyArray(i)
      MyArray(i) = MyArray(j)
      MyArray(j) = Temp
    End If
  Next i
  
  For i = lower To upper
    oCell = oSheet1.getCellByPosition(MyArray(i),1)
    NumC = oCell.CellAddress.Column
    colIndex = NumC+65
    oArg = Array(colIndex)
	  colArray(i) = oSvc.callFunction( "CHAR", oArg)		
  Next i
  
  For j = 0 To t  	
    Max = -1
    Min = 1000000
    sum = 0
    lTick = GetSystemTicks()
    sum = 0

    For i = lower To upper
      oCellRange = oSheet1.getCellRangeByName(colArray(i) + "1:"+ colArray(i) +(rowSize-1))
      oArg = Array(oCellRange)
      sum =  sum + oSvc.callFunction( "COUNT", oArg)
    Next i
    
    lTick = (GetSystemTicks() - lTick)
       
    totalTime = totalTime + lTick
    If lTick > Max Then
    Max = lTick
    End If
    If lTick < Min Then
    Min = lTick
    End If
  Next j
   
  totalTime = totalTime - Max - Min
  oSheet1.getCellByPosition(32, rowIndex).String = totalTime/8
End Sub

'Runs experiments on all spreadsheets specified by  [minRows, maxRows] with stepSize increments.
'This is the main function to be called for running the experiment.

Sub Main
   rowIndex = 1
   For i = 100000 To 500001 Step 200000
	srangeAccess(rowIndex, i)
	randomColumnAccess(rowIndex, i)
	rowIndex=rowIndex+1
   Next i
End Sub
