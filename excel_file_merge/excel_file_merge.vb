Sub MergeFiles()
     
    Dim Path            As String
    Dim FileName        As String
    Dim Wkb             As Workbook
    Dim WS              As Worksheet
     
    Application.EnableEvents = False
    Application.ScreenUpdating = False
     Path = ThisWorkbook.Path & "\September_2016" 'file path subject to change
    FileName = Dir(Path & "\*.xls", vbNormal)
    Do Until FileName = ""
          Set Wkb = Workbooks.Open(FileName:=Path & "\" & FileName, Password:="100") 'Excel file password
        For Each WS In Wkb.Worksheets
            WS.Unprotect Password:="139"
            
            WS.AutoFilterMode = False
            
            'On Error Resume Next
            'WS.ShowAllData
            'On Error GoTo 0
            
            WS.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Next WS
        Wkb.Close False
        FileName = Dir()
    Loop
    
    Dim J As Integer

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add
    Sheets(1).Name = "ComboMerge"
    
    Sheets(3).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    For J = 3 To Sheets.Count
        Sheets(J).Activate
        Range("A1").Select
        Selection.CurrentRegion.Select
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
    
    For Each WS In Sheets
    Application.DisplayAlerts = False
    If WS.Name <> "ComboMerge" Then WS.Delete
    Next
    Application.DisplayAlerts = True
     
End Sub

