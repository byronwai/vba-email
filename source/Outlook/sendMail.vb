Option Explicit
Sub sendMail()
'   https://docs.microsoft.com/en-us/office/vba/excel/concepts/working-with-other-applications/sending-email-to-a-list-of-recipients-using-excel-and-outlook

'   Setting up the Excel variables.
Dim olApp As Object
Dim olMailItm As Object
Dim iCounter As Integer
    
Dim count As Integer
count = 0
Do While isAddr(Cells(count + 2, 1).Text)
'   Create the Outlook application and the empty email.
    Set olApp = CreateObject("Outlook.Application")
    Set olMailItm = olApp.CreateItem(0)
    With olMailItm
       
'   Do additional formatting on the BCC and Subject lines, add the body text from the spreadsheet, and send.
'   Using the email, add multiple recipients, using a list of addresses in column A.
        .To = Cells(count + 2, 1).Text
        .BCC = ""
        .CC = ""
        .Subject = Cells(count + 2, 2).Text
        .Body = Cells(count + 2, 4).Text
        If (Cells(count + 2, 3).Text) <> "" And Dir(Cells(count + 2, 3).Text) <> "" Then
            .Attachments.Add (Cells(count + 2, 3).Text)
        End If
        .Send
    End With
    
'   Clean up the Outlook application.
    Set olMailItm = Nothing
    Set olApp = Nothing
    
    count = count + 1
    
Error_Handling:
    If Err.Description <> "" Then MsgBox Err.Description
   
Loop
End Sub

