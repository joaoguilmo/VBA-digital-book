VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   9195.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label13_Click()

    Unload UserForm3
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

If UserForm2.ListBox1.ListCount <> 1 And UserForm2.ListBox1.ListIndex <> -1 Then

   
    Label10.Caption = UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 2)
    Label11.Caption = UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 5)
    Label12.Caption = UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 1)
    Label6.Caption = UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 3)
    Label7.Caption = UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 4)

End If


End Sub
