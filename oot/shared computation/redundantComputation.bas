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
    oSheet1.getCellByPosition(27,0).String = "Single time"
    oSheet1.getCellByPosition(28,0).String = "Multi Time"
    
  	condition = "1"
	oActiveRange = UsedRange(oDoc.getCurrentController().getActiveSheet())
    oCellRange = oSheet1.getCellRangeByName("J1:J"+rowSize)
    oArg = Array(oCellRange, condition)

   
    totalTimeOne = 0
    For j = 0 To 9  	
        Max = -1
        Min = 1000000
        
        MaxOne = -1
        MinOne = 1000000
        For i = 0 To 4
   
	        lTick = GetSystemTicks()
	 
	       
	       	TotalCount = oSvc.callFunction( "COUNTIFS", oArg)
	        
	       
		        
		    lTick = (GetSystemTicks() - lTick)
		       
		    totalTime = totalTime + lTick
		         
	         If lTick > Max Then
	           Max = lTick
	         End If
	         If lTick < Min Then
	           Min = lTick
	         End If
		     if i = 0 then
		         totalTimeOne = totalTimeOne + lTick
		         
		         If lTick > MaxOne Then
		           MaxOne = lTick
		         End If
		         If lTick < MinOne Then
		           MinOne = lTick
		         End If
		     End If
		         
		   
	    Next i
		
        
        		
   
    Next j
    
    totalTime = totalTime - Max - Min
        totalTimeOne = totalTimeOne - MaxOne - MinOne
   
    oSheet1.getCellByPosition(26,rowIndex).String = rowSize
    oSheet1.getCellByPosition(27, rowIndex).String = totalTimeOne/8
    oSheet1.getCellByPosition(28, rowIndex).String = totalTime/8
   


End Sub



Sub countif_shared
   
   
    Dim rowArray as Variant
	
	j = 1
	For i = 10000 To 500001 Step 10000
		calculateRunTime(j, i)
		j=j+1
	Next i
   


   
   
End Sub
