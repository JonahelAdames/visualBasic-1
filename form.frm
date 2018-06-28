VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4896
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11064
   OleObjectBlob   =   "form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calc_Click()
    If Val(n1) > 100 Or Val(n1) <= 0 Then Label = "Error"
    If Val(n1) <= 100 And Val(n1) >= 90 Then Label = "A"
    If Val(n1) <= 89 And Val(n1) >= 80 Then Label = "B"
    If Val(n1) <= 79 And Val(n1) >= 70 Then Label = "C"
    If Val(n1) <= 69 And Val(n1) >= 65 Then Label = "D"
    If Val(n1) <= 64 And Val(n1) >= 60 Then Label = "E"
    If Val(n1) <= 59 And Val(n1) > 0 Then Label = "F"
End Sub
