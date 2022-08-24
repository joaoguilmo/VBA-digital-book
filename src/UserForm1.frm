VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Caderno Posto de Produção"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub Label5_Click()
    
   
    
    Salvar.Connect
    Salvar.Closeconnection
    MsgBox ("Cadastro realizado com sucesso!")
    
    

End Sub

Private Sub Label6_Click()


Connect_consultar
Closeconnection_consultar
UserForm2.Show


    
End Sub

Private Sub Label8_Click()

    Unload Me

End Sub

Private Sub UserForm_Initialize()

With ComboBox1
  
    .AddItem "POSTO 1"
    .AddItem "POSTO 2"
    .AddItem "POSTO 3"
    .AddItem "POSTO 4"
End With


End Sub
