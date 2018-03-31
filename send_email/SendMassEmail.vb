'Here's a simple, yet powerfull VBA script to send individual emails with attachment

Sub SendMassEmail()

Dim OutlookApp As Object
Dim OutlookMessage As Object
Dim i As Integer

'Optimize Code
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.DisplayAlerts = False

With ActiveSheet
    Lastrow = ActiveSheet.Cells(.Rows.Count, "A").End(xlUp).Row
End With

For i = 2 To Lastrow
  
  Set OutlookApp = CreateObject(class:="Outlook.Application") 'Create Instance of Outlook
  Set OutlookMessage = OutlookApp.CreateItem(0) 'Create a new email message
  
'Create Outlook email with attachment
  On Error Resume Next
    With OutlookMessage
     .To = Cells(i, 2).Value 'email recipient
     .CC = Cells(i, 3).Value 'email copy recipient
     .BCC = ""
     .Subject = Cells(i, 4).Value 'email subject
     .Body = Cells(i, 5).Value 'email body
     .Attachments.Add Cells(i, 6).Value 'email attachment
     .Display
    End With
  On Error GoTo 0
  
Next i

'Clear Memory
  Set OutlookMessage = Nothing
  Set OutlookApp = Nothing
  
'Optimize Code
ExitSub:
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.DisplayAlerts = True
  
MsgBox "Done"

End Sub
