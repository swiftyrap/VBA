Attribute VB_Name = "newEmail"
'author : eddie chung
'date: 23/2/2017
'email: eddiecyc11@gmail.com
'decription:

'1. click tools> references> check "microsoft outlook 1x.0 object library"
'2. create a sheet called "main"
'3. type "product ID" in "a1", type "Style ID" in "b1"
'4. input product id and style id in "a2" & "b2"
'5. create a sheet called "email list"
'6. type "email list" on "a1"
'7. input email address in "a2" & "a3"
'8. create a sheet called "email content"
'9. input subject line in "a1"
'10. input email content on "b1"

Public olApp As Outlook.Application
Public olMail As Outlook.MailItem
Public myCcRecipient As Outlook.Recipient
Public myToRecipient As Outlook.Recipient


Sub newEmailButton()
'new email marco
'input email content, subject line 1 and 2, flag important email as varable in new email method.
Call newEmail(Sheets("email content").Range("b1"), Sheets("email content").Range("a1"), , True)
End Sub

Public Function newEmail(Optional strBodyStr As String, Optional subjectStr As String, Optional subjectStr2 As String, Optional importantFlag As Boolean) As Object
'new email method
Dim strSignature As String

Set olApp = CreateObject("Outlook.Application")
Set olMail = olApp.CreateItem(0)

With olMail

.SentOnBehalfOfName = Sheets("email list").Range("a3").Value 'Sender email

Set myToRecipient = .Recipients.Add(Sheets("email list").Range("a2"))
myToRecipient.Type = olTo
myToRecipient.Resolve

Set myCcRecipient = .Recipients.Add(Sheets("email list").Range("a3"))
myCcRecipient.Type = olCC
myCcRecipient.Resolve

.BodyFormat = olFormatHTML
.Display

If importantFlag = True Then
    .Importance = 2
End If
.Subject = "Product# " & Sheets("main").Range("a2") & " style# " & Trim(Sheets("main").Range("b2")) & "/ " & subjectStr & " " & subjectStr2
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
