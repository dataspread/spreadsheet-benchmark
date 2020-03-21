REM  *****  BASIC  *****

Function putOverlappingFormula(numOfFormula As Long, oSheet1 As Object)
	for i = 0 to 9
	    oSheet1.getCellRangeByName("U1:U10000").ClearContents(2 ^ i)
	Next i
	For i = 1 To numOfFormula
	    oSheet1.getCellRangeByName("U"&i).Formula = "=SUM(J$2:J$" & (i+1) & ")"
	Next i
End Function

Function putCumulativeFormula(numOfFormula As Long, oSheet1 As Object)
	
	for i = 0 to 9
	    oSheet1.getCellRangeByName("U1:U10000").ClearContents(2 ^ i)
	Next i

	oSheet1.getCellRangeByName("U"&1).Formula = "=SUM(J$2:J$" & (2) & ")"
	For i = 2 To numOfFormula
	    oSheet1.getCellRangeByName("U"&i).Formula = "=SUM(U" & (i-1) & "; J" & (i+1) & ")"	
	Next i
End Function

Sub calculateRepeatedComputationTime(rowIndex As Long, numOfFormula As Long)
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
	oSheet1.getCellByPosition(26,0).String = "number of formula"
	oSheet1.getCellByPosition(27,0).String = "common time"
	oSheet1.getCellByPosition(28,0).String = "shared time"
	
	putOverlappingFormula(numOfFormula, oSheet1)


	For j = 0 To 9
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
	oSheet1.getCellByPosition(26,rowIndex).String = numOfFormula
End Sub

Sub calculateReusableComputationTime(rowIndex As Long, numOfFormula As Long)
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
	oSheet1.getCellByPosition(26,0).String = "number of formula"
	oSheet1.getCellByPosition(27,0).String = "common time"
	oSheet1.getCellByPosition(28,0).String = "shared time"
	
	putCumulativeFormula(numOfFormula, oSheet1)


	For j = 0 To 9
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
	oSheet1.getCellByPosition(28,rowIndex).String = totalTime/8
	oSheet1.getCellByPosition(26,rowIndex).String = numOfFormula
End Sub


Sub main	
	ThisComponent.isAutomaticCalculationEnabled = False

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
	oSheet1.getCellByPosition(26,0).String = "number of formula"
	oSheet1.getCellByPosition(27,0).String = "common time"
	oSheet1.getCellByPosition(28,0).String = "shared time"

	rowIndex = 2
	For i = 10000 to 100001 Step 10000
	    calculateRepeatedComputationTime rowIndex, i
	    rowIndex = rowIndex + 1 
	Next i

        rowIndex = 2
	For i = 10000 to 100001 Step 10000
	    calculateReusableComputationTime rowIndex, i
	    rowIndex = rowIndex + 1 
	Next i
	
	
End Sub
