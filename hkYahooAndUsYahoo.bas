Attribute VB_Name = "hkYahooAndUsYahoo"
'author : eddie chung
'date: 23 Feb 2017
'email: eddiecyc11@gmail.com
'description:
'make sure go to firefox setting and checked "Open new windows in a new tab instead" to open multiple tabs on firefox

Sub hkYahooAndUsYahooFirefoxButton()
Dim i As Long
Dim urlString As String
Dim strSitePath As String

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

strSitePath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
urlString = "hk.yahoo.com"
Shell (strSitePath & " -url " & urlString)

strSitePath = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
urlString = "us.yahoo.com"
Shell (strSitePath & " -url " & urlString)

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Sub
