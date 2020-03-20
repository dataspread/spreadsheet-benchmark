REM  *****  BASIC  *****

Sub SortCells( oShet as Object, rowsize As Long, sortAscending As Boolean)
   
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
   sortDesc(1).Name = "HasHeader" ' Does nothing
   sortDesc(1).Value = false
   
   ' Sort the range.
   OShet.getCellRangeByPosition(0,1,15,rowsize).Sort(sortDesc())
End Sub 


Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
    Dim OldDoc As Object
    Dim OldSheet As Object
    Dim oRange As Object
 
   
    Dim totalTime As Long
    Dim Max As Long
    Dim Min As Long
   
   
    Max = -1
    Min = 1000000
    totalTime = 0
   
    'url = ConvertToURL("~/Downloads/result1/weather" & (rowSize) & ".ods")
    url = ConvertToURL("~/Desktop/Research2018Fall/spring2020/xls/airbnb_unsorted_" & (rowSize) & "k.ods")
    
      
	dim oShet as object, oCtrl as object
		
	dim oDataRange as object, oFiltre as object
	dim oFilterField(0) As New com.sun.star.sheet.TableFilterField   
		
	Dim sortAscending As Boolean
   

    For j = 0 To 9
      
      OldDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open it
      oShet = OldDoc.sheets(0)
      sortAscending = true
      lTick = GetSystemTicks()
      
	  SortCells (OShet, rowSize, sortAscending)
		
       
      lTick = (GetSystemTicks() - lTick)
      
      
       
      totalTime = totalTime + lTick
         
      If lTick > Max Then
           Max = lTick
      End If
         
      If lTick < Min Then
           Min = lTick
      End If
      sortAscending = false
      
	  SortCells (OShet, rowSize, sortAscending)
		  
	  OldDoc.dispose 'close it
   
    Next j
    
    

    totalTime = totalTime - Max - Min
    
    oSheet.getCellByPosition(0,rowIndex).String = rowSize
    oSheet.getCellByPosition(1, rowIndex).String = totalTime/8


End Sub




Sub main
    Dim oDoc As Object
    Dim oSheet As Object
    Dim j as Long

    oDoc = ThisComponent ' the file where the results are written

    oSheet = oDoc.Sheets(0) 'get Sheet1
   
    oSheet.getCellByPosition(0,0).String = "Import Size"
   
    oSheet.getCellByPosition(1, 0).String = "Time (ms)"

    rowIndex = 1 'row id where the current result will be written
   
    For i = 10000 to 500001 Step 10000
        calculateRunTime(oDoc,oSheet,rowIndex,i)
        rowIndex = rowIndex + 1   
    Next i

End Sub
