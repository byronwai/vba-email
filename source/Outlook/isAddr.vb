Option Explicit
Function isAddr(str As String) As Boolean
'   Using regular expression to check if the cell contains an valid address
'   Enable "Microsoft VBScript Regular Expressions 5.5" beforehand

Rem https://stackoverflow.com/questions/46155/how-to-validate-an-email-address-in-javascript
Rem https://officeguide.cc/excel-vba-regular-expressions-regex-tutorial/

If str = "" Then
    isAddr = False
Else
    Dim regEx As New RegExp
    regEx.Pattern = "[a-z0-9!#$%&'*+=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
    If regEx.Test(str) Then
         isAddr = True
    Else
        isAddr = False
    End If
End If

End Function

