Sub calculateRunTime(rowIndex As Long, rowCount As Long)
    
    Dim names As String
    Dim surnames As String
    Dim n As Long
    Dim document As Object
    Dim sheets as Object
    Dim sheet as Object
    Dim replace As Object
	rowSize = rowCount*1000
    names = "yawwn" 
    surnames = "wohoo"
    document = ThisComponent rem .CurrentController.Frame
    rem sheet = doc.CurrentSelection.Spreadsheet
    sheets = document.getSheets()
    sheet = sheets.getByIndex(0)
    oCellRange = sheet.getCellRangeByName("A1:P"+rowSize)
    replace = sheet.createReplaceDescriptor rem document.createReplaceDescriptor in case of Writer
    rem replace.SearchRegularExpression = True
    
        
    Max = -1
    Min = 1000000
    totalTime = 0
    
    MaxReverse = -1
    MinReverse = 1000000
    totalTimeReverse = 0

   rowSize = rowCount*1000
    
       
    sheet.getCellByPosition(26,0).String = "RowSize"
    sheet.getCellByPosition(27,0).String = "FnR time (ms)"
     sheet.getCellByPosition(27,0).String = "Reverse time (ms)"
   
  
        Max = -1
        Min = 1000000
        
        For i = 0 To 9
   
	        lTick = GetSystemTicks()
	 		replace.SearchString = names
        	replace.ReplaceString = surnames
        	oCellRange.replaceAll(replace)
	       
	        lTick = (GetSystemTicks() - lTick)
		       
		    totalTime = totalTime + lTick
		         
	         If lTick > Max Then
	           Max = lTick
	         End If
	         If lTick < Min Then
	           Min = lTick
	         End If
	         
	         lTick = GetSystemTicks()
	 		replace.SearchString = surnames
        	replace.ReplaceString = names
        	oCellRange.replaceAll(replace)
	       
	        lTick = (GetSystemTicks() - lTick)
		       
		    totalTimeReverse = totalTimeReverse + lTick
		         
	         If lTick > MaxReverse Then
	           MaxReverse = lTick
	         End If
	         If lTick < MinReverse Then
	           MinReverse = lTick
	         End If

	    Next i
		totalTime = totalTime - Max - Min
		totalTimeReverse = totalTimeReverse - MaxReverse - MinReverse

   
    sheet.getCellByPosition(26,rowIndex).String = rowSize
    sheet.getCellByPosition(27, rowIndex).String = totalTime/8
    sheet.getCellByPosition(28, rowIndex).String = totalTimeReverse/8
   


End Sub

Sub calculateRunTimeNE(rowIndex As Long, rowCount As Long)
    
    Dim names As String
    Dim surnames As String
    Dim n As Long
    Dim document As Object
    Dim sheets as Object
    Dim sheet as Object
    Dim replace As Object
	rowSize = rowCount*1000
    names = "yawwawaawannnnwn" 
    surnames = "wowerteryertyrtyhoo"
    document = ThisComponent rem .CurrentController.Frame
    rem sheet = doc.CurrentSelection.Spreadsheet
    sheets = document.getSheets()
    sheet = sheets.getByIndex(0)
    oCellRange = sheet.getCellRangeByName("A1:P"+rowSize)
    replace = sheet.createReplaceDescriptor rem document.createReplaceDescriptor in case of Writer
    rem replace.SearchRegularExpression = True
    
        
    Max = -1
    Min = 1000000
    totalTime = 0
    
    MaxReverse = -1
    MinReverse = 1000000
    totalTimeReverse = 0

   rowSize = rowCount*1000
    
       
    sheet.getCellByPosition(26,0).String = "RowSize"
    sheet.getCellByPosition(29,0).String = "NE time (ms)"
   
  
        Max = -1
        Min = 1000000
        
        For i = 0 To 9
   
	        lTick = GetSystemTicks()
	 		replace.SearchString = names
        	replace.ReplaceString = surnames
        	oCellRange.replaceAll(replace)
	       
	        lTick = (GetSystemTicks() - lTick)
		       
		    totalTime = totalTime + lTick
		         
	         If lTick > Max Then
	           Max = lTick
	         End If
	         If lTick < Min Then
	           Min = lTick
	         End If
	         
	         lTick = GetSystemTicks()
	 		

	    Next i
		totalTime = totalTime - Max - Min
		

   
    sheet.getCellByPosition(26,rowIndex).String = rowSize
    sheet.getCellByPosition(29, rowIndex).String = totalTime/8
    
   


End Sub

Sub fnr
   
   
   
	
	j = 1
	For i = 10 To 501 Step 10
		calculateRunTimeNE(j, i)
		j=j+1
	Next i
   


   
   
End Sub
