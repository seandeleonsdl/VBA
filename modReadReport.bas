Attribute VB_Name = "modReadReport"
Public Sub ReadReport(Directory As String, ReportName As String)
    
    'MsgBox "funcReadReport"
    
    Dim ReportSummary As String
    'ReportSummary
    
    '--------------------------
    ' Find/set report workbook
    '--------------------------
    Dim DirectoryContents() As String
    DirectoryContents = modGeneral.LoopThroughContents(Directory)
    
    Dim File As Variant
    Dim WorkbookName As String
  
    For Each File In DirectoryContents
        If InStr(1, File, ReportName) Then
            WorkbookName = Directory & "\" & File
            Exit For
        End If
    Next
    
    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(WorkbookName)
            
    '--------------------
    ' ItemTypeCount check
    '--------------------
    Dim ItemTypeCount As Integer
    ItemTypeCount = wb.Sheets(1).Range("C2").End(xlDown).Row - 3
    
    For i = 1 To ItemTypeCount
        ItemCode = Int(Left(wb.Sheets(1).Range("B4").Offset(i - 1, 0).Value2, 3))
        TypeCount = wb.Sheets(1).Range("C4").Offset(i - 1, 0).Value2
        'MsgBox ItemCode & " " & TypeCount
                
        If ItemCode = 15 And TypeCount >= 5 Then wb.Sheets(1).Range("C4").Offset(i - 1, 0).Interior.Color = RGB(255, 222, 33)
        If ItemCode = 15 And TypeCount >= 5 Then wb.Sheets(1).Range("C4").Offset(i - 1, 0).Interior.Color = RGB(255, 222, 33)
        If ItemCode = 18 And TypeCount >= 5 Then wb.Sheets(1).Range("C4").Offset(i - 1, 0).Interior.Color = RGB(255, 222, 33)
        If ItemCode = 23 And TypeCount >= 5 Then wb.Sheets(1).Range("C4").Offset(i - 1, 0).Interior.Color = RGB(255, 222, 33)
    Next
            
    '-----------------------
    ' Yield Per Wafer Check
    '-----------------------
    Dim ItemSublot As Integer, MinQty As Double
    ItemSublot = Application.WorksheetFunction.Count(wb.Sheets(1).Range("F:F"))
    MinQty = 100#
    
    Instance = Rows.Count
    For i = 1 To ItemSublot
        Instance = wb.Sheets(1).Range("F" & Instance).End(xlUp).Row
        ItemQty = wb.Sheets(1).Range("F" & Instance).Value
        If MinQty > ItemQty Then MinQty = ItemQty
    Next
            
    If MinQty < 90# Then
        MsgBox "Lowest Qty: " & MinQty & vbNewLine & "Low Qty"
    End If
    
    'wb.Close False
    
    Exit Sub
End Sub


