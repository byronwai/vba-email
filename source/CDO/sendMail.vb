Option Explicit
Sub sendMail()
'   "Microsoft CDO for Windows 2000" maybe required

Rem https://www.makeuseof.com/tag/send-emails-excel-vba/
Rem https://blog.xuite.net/saladoil/excel/8771822-%E4%BB%A5VBA%E5%82%B3%E9%80%81%E9%83%B5%E4%BB%B6--CDO%E7%89%A9%E4%BB%B6

Dim CDO_Mail As Object
Dim CDO_Config As Object
Dim SMTP_Config As Variant

Set CDO_Mail = CreateObject("CDO.Message")
On Error GoTo Error_Handling

Set CDO_Config = CreateObject("CDO.Configuration")
CDO_Config.Load -1

'   Setup SMTP connection
Set SMTP_Config = CDO_Config.Fields

With SMTP_Config
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "<URL of SMTP server>"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "<email address>"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "<password>"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Update
End With
Dim count As Integer
count = 0

'   Loop for Sheet1 starting from Row 2, until Cell in Column A is empty or is an INVALID email address
'   '   Column A: Email address
'   '   Column B: Subject
'   '   Column C: Canonical path of attahment, attach when file is found
'   '   Column D: Text body
Do While isEmail(Cells(count + 2, 1).Text)
    With CDO_Mail
        Set .Configuration = CDO_Config
    End With
    '   Reset and set email content
    CDO_Mail.Attachments.DeleteAll
    CDO_Mail.Subject = Cells(count + 2, 2).Text
    CDO_Mail.From = "<Self Define Name>"
    CDO_Mail.To = Cells(count + 2, 1).Text
    CDO_Mail.CC = ""
    CDO_Mail.BCC = ""
    CDO_Mail.TextBody = Cells(count + 2, 4).Text
    If (Cells(count + 2, 3).Text) <> "" And Dir(Cells(count + 2, 3).Text) <> "" Then
        CDO_Mail.AddAttachment (Cells(count + 2, 3).Text)
    End If
    CDO_Mail.Send

Error_Handling:
    If Err.Description <> "" Then MsgBox Err.Description
    
    count = count + 1
Loop

'   Show number of email sent
MsgBox (count) & " mail has been sent."

Set CDO_Mail = Nothing

End Sub
