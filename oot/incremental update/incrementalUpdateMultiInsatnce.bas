REM  *****  BASIC  *****

Sub putFormula(rowSize As Long, oSheet1 As Object)
	For i = 1 To rowSize
		oSheet1.getCellRangeByName("T"&i).Formula = "=COUNTIF(J$2:J$" & 500000 &";1)"
	Next i
End Sub

Sub calculateRunTime(rowIndex As Long, rowSize As Long)
	Dim myDoc As Object
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

	myDoc = ThisComponent
	oSheet1 = myDoc.Sheets(0)
	oSheet1.getCellByPosition(26,0).String = "number of formulas"
	oSheet1.getCellByPosition(27,0).String = "recalculation"
	
	putFormula(rowSize, oSheet1)

	For j = 0 To 9
	    if oSheet1.getCellByPosition(9,1).Value <> 0 Then
	        oSheet1.getCellByPosition(9,1).Value = 0
	    Else
	        oSheet1.getCellByPosition(9,1).Value = 1
	    Endif
		
	    lTick = GetSystemTicks()

	    ThisComponent.calculateAll()
	
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
	oSheet1.getCellByPosition(27,rowIndex).String = totalTime/8
	oSheet1.getCellByPosition(26,rowIndex).String = rowSize
End Sub

Sub Main
	
	ThisComponent.isAutomaticCalculationEnabled = False
	calculateRunTime(1,1)
	rowIndex = 2
	For i = 100 To 1001 Step 100
            calculateRunTime(rowIndex, i)
            rowIndex=rowIndex+1
	Next i
End Sub
