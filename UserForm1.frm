VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'author: eddie chung
'date: 2 Mar 2017
'email: eddiecyc11@gmail.com
'description:
'1. create sheet "combo box list"
'2. input "country list on a1.
'3. input country on column A.

Private Sub UserForm_Initialize()

Dim i As Long, i2 As Long
Dim myArr() As String, myArr2() As String

With Sheets("Combo Box List")

    For i = 2 To .Range("a1048576").End(xlUp).Row
        ReDim Preserve myArr(2 To i)
        If Not IsEmpty(.Cells(i, "a")) Then
            myArr(i) = .Cells(i, "a")
        End If
    Next i
End With

Me.countryComboBox.List = myArr()

Erase myArr

End Sub

Private Sub countryComboBox_Change()

Me.Label7.Caption = Me.countryComboBox.Value

End Sub

Private Sub TextBox2_Change()

Me.Label9.Caption = Me.TextBox2.Value

End Sub

