REM  *****  BASIC  *****

Function UsedRange(oSheet As Variant) As Variant
    Dim oCursor As Variant
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    oCursor.gotoStartOfUsedArea(True)
    UsedRange = oCursor
End Function

Sub calculateRunTime(rowIndex As Long, rowSize As Long)
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
	

    oSvc = createUnoService( "com.sun.star.sheet.FunctionAccess")
    oDoc = ThisComponent
    oSheet1 = oDoc.Sheets(0)
       
    oSheet1.getCellByPosition(26,0).String = "RowSize"
    oSheet1.getCellByPosition(27,0).String = "sum"
    oSheet1.getCellByPosition(28,0).String = "Time"
    
   
        
    Dim sum As Long
    
    
   	Dim MyArray(rowSize-1) As Long
   	lower = 1
   	upper = rowSize-1
   	For i = lower To upper
   		MyArray(i) = i+1
   	Next i
   
    For j = 0 To 9  	
        Max = -1
        Min = 1000000
        sum = 0
        lTick = GetSystemTicks()
        
        For i = lower To upper
        	sum = sum + oSheet1.getCellByPosition(9,MyArray(i)).Value
        	'print oSheet1.getCellByPosition(10,MyArray(i)).Value
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

	oSheet1.getCellByPosition(26,rowIndex).String = rowSize
    oSheet1.getCellByPosition(27,rowIndex).String = sum
    oSheet1.getCellByPosition(28, rowIndex).String = totalTime/8
   


End Sub

Sub calculateRunTime2(rowIndex As Long, rowSize As Long)
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
	

    oSvc = createUnoService( "com.sun.star.sheet.FunctionAccess")
    oDoc = ThisComponent
    oSheet1 = oDoc.Sheets(0)
       
    oSheet1.getCellByPosition(29,0).String = "sum"
    oSheet1.getCellByPosition(30,0).String = "Rand Time"
    
   
        
    Dim sum As Long
    
    
   	Dim MyArray(rowSize-1) As Long
   	lower = 1
   	upper = rowSize-1
   	For i = lower To upper
   		MyArray(i) = i+1
   	Next i
   	
   	'Dim Temp as Long
   	For i = lower To upper
   		j = CLng((upper - i) * rnd + i)
   		If i <> j Then
            Temp = MyArray(i)
            MyArray(i) = MyArray(j)
            MyArray(j) = Temp
        End If
   	Next i
   
    For j = 0 To 9  	
        Max = -1
        Min = 1000000
        sum = 0
        lTick = GetSystemTicks()
        
        For i = lower To upper
        	sum = sum + oSheet1.getCellByPosition(9,MyArray(i)).Value
        	'print oSheet1.getCellByPosition(10,MyArray(i)).Value
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

	oSheet1.getCellByPosition(26,rowIndex).String = rowSize
    oSheet1.getCellByPosition(29,rowIndex).String = sum
    oSheet1.getCellByPosition(30, rowIndex).String = totalTime/8
   


End Sub


Sub Main
   
    Dim rowArray as Variant
	
	j = 1
	For i = 100000 To 500001 Step 200000
		calculateRunTime(j, i)
		calculateRunTime2(j, i)
		j=j+1
	Next i
   
   
End Sub
