Option Explicit
Sub Button2_Click()
    On Error GoTo ErrHandler
    
    ' https://www.encodedna.com/excel/how-to-parse-outlook-emails-and-show-in-excel-worksheet-using-vba.htm
    ' https://learndataanalysis.org/properly-search-an-email-emails-in-outlook-vba/
    
    ' Set Outlook application object.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    ' Create and Set a NameSpace OBJECT.
    Dim objNSpace As Namespace
    ' The GetNameSpace() method will represent a specified Namespace.
    Set objNSpace = objOutlook.GetNamespace("MAPI")
    
    ' Create a folder object.
    Dim myFolder As MAPIFolder
    Set myFolder = objNSpace.GetDefaultFolder(olFolderInbox)
    
    Dim filtered_items As Items
    Dim strFilter As String
    
    Dim currRow As Integer
    currRow = 2
    Do While IsEmpty(Cells(currRow, 2)) = False
               
        strFilter = "@SQL= urn:schemas:httpmail:sender LIKE '%" + Cells(currRow, 2).Text + "%'"
        Set filtered_items = myFolder.Items.Restrict(strFilter)
        
        If filtered_items.count = 0 Then GoTo ContinueLoop
        
        filtered_items.Sort "ReceivedTime", True
        If filtered_items.GetFirst.Class = olMail Then
        
            Dim objMail As Outlook.MailItem
            Set objMail = filtered_items.GetFirst
    
            Cells(currRow, 3) = objMail.SenderEmailAddress
            Cells(currRow, 4) = objMail.To
            Cells(currRow, 5) = objMail.Subject
            Cells(currRow, 6) = objMail.ReceivedTime
        End If
        
        
ContinueLoop:
        currRow = currRow + 1
    Loop
    
    
    ' Release.
    Set objOutlook = Nothing
    Set objNSpace = Nothing
    Set myFolder = Nothing
    
ErrHandler:
    Debug.Print Err.Description

End Sub

