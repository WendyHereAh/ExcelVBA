Attribute VB_Name = "Module1"
Sub sendEmail()

Application.ScreenUpdating = False
   'Setting up the Excel variables.
   Dim olApp As Object
   Dim iCounter As Integer: Dim Dest As Variant: Dim SDest As String: Dim i As Integer
   Dim emailTemplate As Object: Dim wdeditor As Object
   Dim Path As String: Dim OutAccount As Account
   
   
   'Create the Outlook application and the empty email.
   Set olApp = CreateObject("Outlook.Application")
   LastRow = ActiveSheet.Range("A1").End(xlDown).Row
   Path = ActiveSheet.Cells(2, 2).Value
   
   Set emailTemplate = olApp.CreateItemFromTemplate(Path)
   
   'Use the first account, see that Item is 1 now
   Set OutAccount = olApp.Session.Accounts.Item(3)
   'Using the email, add multiple recipients, using a list of addresses in column A.
   With emailTemplate
       SDest = ""
       For iCounter = 3 To LastRow
        If Last > 302 Then
            Exit For
        Else
            If SDest = "" Then
                SDest = Cells(iCounter, 1).Value
            ElseIf iCounter < 303 Then
                SDest = SDest & ";" & Cells(iCounter, 1).Value
            End If
        End If
       Next iCounter
       
    'Do additional formatting on the BCC and Subject lines, add the body text from the spreadsheet, and send.
       .SendUsingAccount = OutAccount
       .BCC = SDest
       .Subject = ActiveSheet.Cells(2, 1).Value
       .Display
   End With
   If LastRow > 302 Then
        Confirm = MsgBox("300 emails have been sent!", vbYesNo)
        Response = MsgBox("Would you like to delete sent records?", vbOKCancel)
        If Response = vbOK Then
           Call Clear_first_300
        Else: End If
   Else: MsgBox LastRow - 2 & " emails have been sent!"
         Range("A4:B" & LastRow).Select
         Selection.Delete
         MsgBox "Deleted!"
   End If
   
   'Clean up the Outlook application.
   Set emailTemplate = Nothing
   Set olApp = Nothing
End Sub
