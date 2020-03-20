VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Function GetMemUsage(ByVal rowIndex, ByVal rowSize)
  
  ' Returns the current Excel.Application
  ' memory usage in MB
  Dim directory As String
  Dim Name As String
  Dim tmpMemUsage As Long
  Dim memUsage As Long
  Dim minUsage As Long
  Dim maxUsage As Long
  directory = "C:\Users\rzhang74\Desktop\Research2018Fall\Spring2020\memory\"
  Name = "airbnb_" & CStr(rowSize) & ".xlsx"
  Filename = Dir(directory & Name)
    
  minUsage = 10000000
  maxUsage = -1
  For i = 1 To 10
    Workbooks.Open (directory & Filename)
          
    Set objSWbemServices = GetObject("winmgmts:")
    
    tmpMemUsage = objSWbemServices.Get( _
      "Win32_Process.Handle='" & _
      GetCurrentProcessId & "'").WorkingSetSize / 1024
      
    memUsage = memUsage + tmpMemUsage
    Set objSWbemServices = Nothing
    
    Workbooks(Filename).Close
    
    If tmpMemUsage < minUsage Then
        minUsage = tmpMemUsage
    End If
    If tmpMemUsage > maxUsage Then
        maxUsage = tmpMemUsage
    End If
    'Sheets("import").Cells(rowIndex, i + 2).Value = tmpMemUsage
  Next i
  
  Sheets("import").Cells(rowIndex, 1).Value = rowSize
    
  Sheets("import").Cells(rowIndex, 2).Value = (memUsage - maxUsage - minUsage) / 1024 / 8
  
End Function

Private Function GetMemUsageFormula(ByVal rowIndex, ByVal rowSize)
  
  ' Returns the current Excel.Application
  ' memory usage in MB
  Dim directory As String
  Dim Name As String
  Dim tmpMemUsage As Long
  Dim memUsage As Long
  Dim minUsage As Long
  Dim maxUsage As Long
  directory = "C:\Users\rzhang74\Desktop\Research2018Fall\Spring2020\memory\formula\"
  Name = "weather" & CStr(rowSize) & ".xlsx"
  Filename = Dir(directory & Name)
    
  minUsage = 10000000
  maxUsage = -1
  For i = 1 To 10
    Workbooks.Open (directory & Filename)
          
    Set objSWbemServices = GetObject("winmgmts:")
    
    tmpMemUsage = objSWbemServices.Get( _
      "Win32_Process.Handle='" & _
      GetCurrentProcessId & "'").WorkingSetSize / 1024
      
    memUsage = memUsage + tmpMemUsage
    Set objSWbemServices = Nothing
    
    Workbooks(Filename).Close
    
    If tmpMemUsage < minUsage Then
        minUsage = tmpMemUsage
    End If
    If tmpMemUsage > maxUsage Then
        maxUsage = tmpMemUsage
    End If
    'Sheets("import").Cells(rowIndex, i + 2).Value = tmpMemUsage
  Next i
  
  'Sheets("import").Cells(rowIndex, 1).Value = rowSize
    
  Sheets("import").Cells(rowIndex, 3).Value = (memUsage - maxUsage - minUsage) / 1024 / 8
  
End Function

Private Function GetMemUsagePivot(ByVal rowIndex, ByVal rowSize)
  
  ' Returns the current Excel.Application
  ' memory usage in MB
  Dim directory As String
  Dim Name As String
  Dim tmpMemUsage As Long
  Dim memUsage As Long
  Dim minUsage As Long
  Dim maxUsage As Long
  directory = "C:\Users\rzhang74\Desktop\Research2018Fall\Spring2020\memory\pivot\"
  Name = "weather" & CStr(rowSize) & ".xlsx"
  Filename = Dir(directory & Name)
    
  minUsage = 10000000
  maxUsage = -1
  For i = 1 To 10
    Workbooks.Open (directory & Filename)
          
    Set objSWbemServices = GetObject("winmgmts:")
    
    tmpMemUsage = objSWbemServices.Get( _
      "Win32_Process.Handle='" & _
      GetCurrentProcessId & "'").WorkingSetSize / 1024
      
    memUsage = memUsage + tmpMemUsage
    Set objSWbemServices = Nothing
    
    Workbooks(Filename).Close
    
    If tmpMemUsage < minUsage Then
        minUsage = tmpMemUsage
    End If
    If tmpMemUsage > maxUsage Then
        maxUsage = tmpMemUsage
    End If
    'Sheets("import").Cells(rowIndex, i + 2).Value = tmpMemUsage
  Next i
  
  'Sheets("import").Cells(rowIndex, 1).Value = rowSize
    
  Sheets("import").Cells(rowIndex, 4).Value = (memUsage - maxUsage - minUsage) / 1024 / 8
  
End Function

Private Function GetMemUsageExcel(ByVal rowIndex, ByVal rowSize)
  
  ' Returns the current Excel.Application
  ' memory usage in MB
  Dim directory As String
  Dim Name As String
  Dim tmpMemUsage As Long
  Dim memUsage As Long
  Dim minUsage As Long
  Dim maxUsage As Long
  
  minUsage = 10000000
  maxUsage = -1
  For i = 1 To 10
    Set objSWbemServices = GetObject("winmgmts:")
    
    tmpMemUsage = objSWbemServices.Get( _
      "Win32_Process.Handle='" & _
      GetCurrentProcessId & "'").WorkingSetSize / 1024
      
    memUsage = memUsage + tmpMemUsage
    Set objSWbemServices = Nothing
    
    If tmpMemUsage < minUsage Then
        minUsage = tmpMemUsage
    End If
    If tmpMemUsage > maxUsage Then
        maxUsage = tmpMemUsage
    End If
    'Sheets("import").Cells(rowIndex, i + 2).Value = tmpMemUsage
  Next i
  
  'Sheets("import").Cells(rowIndex, 1).Value = rowSize
    
  Sheets("import").Cells(rowIndex, 5).Value = (memUsage - maxUsage - minUsage) / 1024 / 8
  
End Function

Sub Test1()

'Declare row iterator i

    Dim i As Long

'set headers at the first row. 17,18, and 19 are the coulmn numbers
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("import").Cells(1, 1).Value = "Rows Imported"
    Sheets("import").Cells(1, 2).Value = "Value Size (MB)"
    Sheets("import").Cells(1, 3).Value = "Formula Size (MB)"
    Sheets("import").Cells(1, 4).Value = "Pivot Size (MB)"
    
    i = 2
    For j = 100000 To 500001 Step 100000
        GetMemUsage i, j
        
        'GetMemUsageFormula i, j
        
        'GetMemUsagePivot i, j
        i = i + 1
          
    Next j
    
    'GetMemUsageExcel 2, 0
    Application.ScreenUpdating = True
End Sub





