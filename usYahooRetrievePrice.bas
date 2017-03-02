Attribute VB_Name = "usYahooRetrievePrice"
'author: eddie chung
'date: 2 Mar 2017
'email: eddiecyc11@gmail.com
'description:
'1. create sheets named "main".
'2. input "stock name" on a1.
'3. input " current stock price" on b1.
'4. input " previous closing price" on c1.
'5. input "market open price" on d1.
'6. input stock name that you would like to check on column A. e.g. google, apple.


Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub usYahooRetrievePriceButton()

Dim objIE As Object
Dim urlString As String
Dim htmlDoc As HTMLDocument

Dim i As Long

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

For i = 2 To Sheets("main").Range("a1048576").End(xlUp).Row
    If Not IsEmpty(Sheets("main").Cells(i, "a")) Then

Set objIE = New InternetExplorer
urlString = "https://us.yahoo.com"

Sleep (5000)
objIE.Navigate urlString
objIE.Visible = 1
Do Until objIE.ReadyState = 4
DoEvents
Loop
Set htmlDoc = objIE.Document

Sleep (5000)

htmlDoc.getElementsByClassName("D(ib) Mstart(21px) Mend(13px)")(2).Children(0).Click

Sleep (5000)
htmlDoc.all("p").Value = Trim(Sheets("main").Cells(i, "a"))
Sleep (8000)
htmlDoc.getElementById("search-buttons").Children(0).Click
Sleep (8000)
       Sheets("main").Cells(i, "b") = htmlDoc.getElementsByClassName("Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)")(0).innerText
       Sheets("main").Cells(i, "c") = htmlDoc.getElementsByClassName("Ta(end) Fw(b)")(0).innerText
       Sheets("main").Cells(i, "d") = htmlDoc.getElementsByClassName("Ta(end) Fw(b)")(1).innerText
Sleep (5000)

Set htmlDoc = Nothing
objIE.Quit
Set objIE = Nothing
    End If
Next i

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic


Set objIE = Nothing
Set htmlDoc = Nothing
objIE.Exit

End Sub
