Function Log(rngLastRow As Long, strLogInfo As String, strLog1 As String, strLog2 As String, strLog3 As String, strLog4 As String, strLog5 As String, strLog6 As String, strLog7 As String, strLog8 As Double, Optional strsheetName As String) As String
    Dim ws As Worksheet
    If strsheetName = "" Then strsheetName = "Dump Truck 5 Year by Half Dispo"
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(strsheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Worksheet '" & strsheetName & "' not found.", vbExclamation
        Log = "Error: Worksheet '" & strsheetName & "' not found."
        Exit Function
    End If
    
    
    
    With ws
        .Cells(rngLastRow, 1).Value = strLogInfo
        .Cells(rngLastRow, 2).Value = strLog1
        .Cells(rngLastRow, 3).Value = strLog2
        .Cells(rngLastRow, 4).Value = strLog3
        .Cells(rngLastRow, 5).Value = strLog4
        .Cells(rngLastRow, 6).Value = strLog5
        .Cells(rngLastRow, 7).Value = strLog6
        .Cells(rngLastRow, 8).Value = strLog7
        .Cells(rngLastRow, 9).Value = strLog8
     
    End With
    
    Log = "Date=" & Date & " Time=" & Time & " Message=" & strLogInfo
End Function

Function DeleteLog(rngLastRow As Long, Optional strsheetName As String)

    Dim ws As Worksheet
    If strsheetName = "" Then strsheetName = "Dump Truck 5 Year by Half Dispo"
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(strsheetName)
    
  
    
    With ws
        .Cells(rngLastRow, 1).Value = ""
        .Cells(rngLastRow, 2).Value = ""
        .Cells(rngLastRow, 3).Value = ""
        .Cells(rngLastRow, 4).Value = ""
        .Cells(rngLastRow, 5).Value = ""
        .Cells(rngLastRow, 6).Value = ""
        .Cells(rngLastRow, 7).Value = ""
        .Cells(rngLastRow, 8).Value = ""
        .Cells(rngLastRow, 9).Value = ""
      
    End With
    
    

End Function

Sub EntryButton1_Click()

    Dim strEQID As String
    
    Dim strYear As String
    
    Dim strMake As String
    
    Dim strModel As String
    
    Dim strVIN As String
    
    Dim strDesc As String
    
    Dim strCategory As String
    
    Dim strCatDesc As String
    
    Dim dblMiles As Double
    
    Dim rngLastRow As Long
    
    rngLastRow = CLng(InputBox("Enter which row you would like to replace:"))
   
    strEQID = InputBox("Enter equipment ID:")
    
    strYear = InputBox("Enter equipment year:")
    
    strMake = InputBox("Enter the make of the equipment:")
    
    strModel = InputBox("Enter the model of the equipment:")
    
    strVIN = InputBox("Enter the VIN Number:")
    
    strDesc = InputBox("Enter the description:")
    
    strCategory = InputBox("Enter the category (TRK.DUMP):")
    
    strCatDesc = InputBox("Enter the category description:")
    
    dblMiles = Val(InputBox("Enter the current mileage of the equipment"))
    
   
    

    Call Log(rngLastRow, strEQID, strYear, strMake, strModel, strVIN, strDesc, strCategory, strCatDesc, dblMiles, "")

End Sub

Sub DeleteButton1_Click()

    Dim rngLastRow As Long
    
    rngLastRow = CLng(InputBox("Enter the row number you want to remove:"))
    
    Call DeleteLog(rngLastRow, "Dump Truck 5 Year by Half Dispo")
    

End Sub
