Attribute VB_Name = "replyAllEmail"
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

Sub replyAllEmailButton()

'reply all email marco

Call replyAllEmail(Sheets("email content").Range("b1"), Sheets("email list").Range("a2"), Sheets("email list").Range("a3")) 'call replyAllEmail function
End Sub
Public Function replyAllEmail(Optional strBodyStr As String, Optional SentOnBehalfOfNameStr As String, Optional myCcRecipientStr As String) As Object

'reply all email function

Dim strSignature As String

 'click on an email to trigger the GetcurrentItem() function
If Not GetCurrentItem() Is Nothing Then
    Set olMail = GetCurrentItem().ReplyAll
End If
    
With olMail

Set myAllRecipients = .Recipients

For Each myRecipient In myAllRecipients

If myRecipient.AddressEntry.AddressEntryUserType = olExchangePublicFolderAddressEntry And myRecipient.AddressEntry.Name = Sheets("Email List").Range("a2") Then
myRecipient.Delete
End If

Next myRecipient

.SentOnBehalfOfName = SentOnBehalfOfNameStr 'Sender email address

Set myCcRecipient = .Recipients.Add(myCcRecipientStr) 'cc email address
myCcRecipient.Type = olCC
myCcRecipient.Resolve

.BodyFormat = olFormatHTML
 'Show signature and content in html format
strSignature = .HTMLBody
.HTMLBody = strBodyStr & strSignature
.Display

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

