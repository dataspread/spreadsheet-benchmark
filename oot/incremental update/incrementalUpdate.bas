REM  *****  BASIC  *****

Sub putFormula(rowSize As Long, oSheet1 As Object)
	For i = 1 To 1000
		
		oSheet1.getCellRangeByName("T"&i).Formula = "=SUM(J$2:J$" & rowSize  &";1)"
	
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
	oSheet1.getCellByPosition(26,0).String = "Row Size"
	oSheet1.getCellByPosition(27,0).String = "materialization"
	
	totalTime = 0
	Min = 9999
	Max = 0
	
	putFormula(rowSize, oSheet1)

	For j = 0 To 9
		'myDoc.addActionLock()
		' --- modify your cells here ---

		if oSheet1.getCellByPosition(9,1).Value <> 0 Then
			oSheet1.getCellByPosition(9,1).Value = 0
		Else
			oSheet1.getCellByPosition(9,1).Value = 1
		Endif
		
		lTick = GetSystemTicks()
		
		'myDoc.removeActionLock()
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
	j=1
	For i = 10000 To 10001 Step 10000
		calculateRunTime(j, i)
		j=j+1
	Next i
	
End Sub
