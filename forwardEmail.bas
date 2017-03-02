Attribute VB_Name = "forwardEmail"
'author : eddie chung
'date: 23 Feb 2017
'email: eddiecyc11@gmail.com
'decription:
'1. click tools> references> check "microsoft outlook 1x.0 object library"
'2. create a sheet called "main"
'3. create a sheet called "email list"
'4. type "email list" on "a1"
'5. input email address in "a2" & "a3"
'6. create a sheet called "email content"
'7. input email content on "b1"

Sub forwardEmailButton()
'forward email marco
'input email body, sender email address, to email address. cc email address in bracket.
Call forwardEmail(Sheets("email content").Range("b1"), Sheets("email list").Range("a3"), Sheets("email list").Range("a2"), Sheets("email list").Range("a3")) 'call forwardEmail function
End Sub


Public Function forwardEmail(Optional strBodyStr As String, Optional SentOnBehalfOfNameStr As String, Optional myToRecipientStr As String, Optional myCcRecipientStr As String) As Object

'forward email function

Dim strSignature As String

 'click on an email to trigger the GetcurrentItem() function
 
If Not GetCurrentItem() Is Nothing Then
    Set olMail = GetCurrentItem().Forward
End If
    
With olMail

.SentOnBehalfOfName = SentOnBehalfOfNameStr 'Sender email address

Set myToRecipient = .Recipients.Add(myToRecipientStr) 'to email address
myToRecipient.Type = olTo
myToRecipient.Resolve

Set myCcRecipient = .Recipients.Add(myCcRecipientStr) 'cc email address
myCcRecipient.Type = olCC
myCcRecipient.Resolve

.BodyFormat = olFormatHTML
strSignature = .HTMLBody
.HTMLBody = strBodyStr & strSignature
.Display 'Show signature and content in html format

olMail.UnRead = False

End With
    
Set myCcRecipient = Nothing
Set olMail = Nothing
Set olApp = Nothing

End Function

Public Function GetCurrentItem() As Object

'GetCurrentItem function

Set olApp = CreateObject("Outlook.Application")

On Error Resume Next
Select Case TypeName(olApp.ActiveWindow)
    Case "Explorer"
        Set GetCurrentItem = olApp.ActiveExplorer.Selection.Item(1)
    Case "Inspector"
        Set GetCurrentItem = olApp.ActiveInspector.CurrentItem
End Select

End Function


