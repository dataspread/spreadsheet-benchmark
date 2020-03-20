REM  *****  BASIC  *****

Function MakePropertyValue( Optional cName As String, Optional uValue ) As com.sun.star.beans.PropertyValue
   Dim oPropertyValue As New com.sun.star.beans.PropertyValue
   If Not IsMissing( cName ) Then
      oPropertyValue.Name = cName
   EndIf
   If Not IsMissing( uValue ) Then
      oPropertyValue.Value = uValue
   EndIf
   MakePropertyValue() = oPropertyValue
End Function


Sub calculateRunTime(oDoc as Object, oSheet as Object, rowIndex As Long, rowSize As Long)
    Dim OldDoc As Object
    Dim OldSheet As Object
    Dim oRange As Object
 
   
    Dim totalTime As Long
    Dim Max As Long
    Dim Min As Long
   

    totalTime = 0
   
    url = ConvertToURL("~/Desktop/Research2018Fall/spring2020/open/airbnb_" & (rowSize) & "k.ods")
    rowCount = rowSize*1000
    OldDoc = StarDesktop.loadComponentFromURL(url,"_Blank",0,Array()) 'open it
      

    For j = 0 To 9
     
     
      lTick = GetSystemTicks()
      
      temp = OldDoc.sheets(0).getCellByPosition(9, 1).ConditionalFormat
      oldCondFormat = OldDoc.sheets(0).getCellByPosition(9, 1).ConditionalFormat
      oldCondFormat.addNew( Array( _
        MakePropertyValue( "Operator", com.sun.star.sheet.ConditionOperator.FORMULA ),_
        MakePropertyValue( "Formula1", "J2=1" ), _
        MakePropertyValue( "SourcePosition", OldDoc.sheets(0).getCellByPosition(9, 1).getCellAddress() ), _
        MakePropertyValue( "StyleName", "Good" ))) 
      OldDoc.sheets(0).getCellRangeByName("J2:J"&rowCount).setPropertyValue("ConditionalFormat", oldCondFormat)
    
       
      lTick = (GetSystemTicks() - lTick)
      
      OldDoc.sheets(0).getCellRangeByName("J2:J"&rowCount).setPropertyValue("ConditionalFormat", temp)
      
      totalTime = totalTime + lTick
         
      If lTick > Max Then
           Max = lTick
      End If
         
      If lTick < Min Then
           Min = lTick
      End If
        

   
    Next j
    
    OldDoc.dispose 'close it

    totalTime = totalTime - Max - Min
    
    oSheet.getCellByPosition(0,rowIndex).String = rowCount
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
