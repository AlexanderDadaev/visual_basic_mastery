Sub SaveAs_macro()

Dim IndPronik As String
Dim IndPronik_otp As String

IndPronik = "Individ_Proniknovenie.xlsx"
IndPronik_otp = "Individ_Proniknovenie_otp.xlsx"

Application.DisplayAlerts = False
ThisWorkbook.CheckCompatibility = False

ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "_otp.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = False
        Sheets("SheetB").Delete
        Sheets("SheetC").Delete
        
Workbooks.Open Filename:=ThisWorkbook.Path & "\SaveAs_Meta.xlsm"
ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & IndPronik, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = False
        Sheets("SheetA").Delete
        Sheets("SheetC").Delete

Workbooks.Open Filename:=ThisWorkbook.Path & "\SaveAs_Meta.xlsm"
ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & IndPronik_otp, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = False
        Sheets("SheetA").Delete
        Sheets("SheetB").Delete

ActiveWorkbook.Close SaveChanges:=True
ActiveWorkbook.Close SaveChanges:=True
ActiveWorkbook.Close SaveChanges:=True

'Application.DisplayAlerts = False

'    Sheets("SheetB").Delete
 '   Sheets("SheetC").Delete

'ActiveWorkbook.Close SaveChanges:=True

'Invidiv-pronik
'Workbooks.Open Filename:=ThisWorkbook.Path & "\SaveAs_Meta.xlsm"
'ThisWorkbook.Activate

End Sub
