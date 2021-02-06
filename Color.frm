VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSample20 
   Caption         =   "Color"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   OleObjectBlob   =   "Color.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmSample20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If txtNew = "" Then Exit Sub
    listColor.AddItem txtNew
    listColor.ListIndex = listColor.ListCount - 1
    txtNew = "": txtNew.SetFocus
End Sub

Private Sub cmdRemove_Click()
    If listColor.ListIndex = -1 Then Exit Sub
    listColor.RemoveItem listColor.ListIndex
    listColor.ListIndex = -1
End Sub

Private Sub UserForm_Initialize()
    With listColor
        .AddItem "White"
        .AddItem "Black"
        .AddItem "Red"
        .AddItem "Yellow"
        .AddItem "Pink"
        .AddItem "Green"
        .AddItem "Blue"
    End With
End Sub

