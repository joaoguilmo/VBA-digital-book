VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Lista defeitos e falhas "
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub cabecalho()

    With ListBox1
    
        .AddItem
        .List(0, 0) = "Codigo"
        .List(0, 1) = "Data"
        .List(0, 2) = "Posto"
        .List(0, 3) = "Equipamento"
        .List(0, 4) = "Descrição"
        .List(0, 5) = "Nota de Manutenção"
        .List(0, 6) = "Corrigido"
        
    End With
    

End Sub


Private Sub ComboBox1_Change()

ListBox1.Clear
cabecalho

Connect_consultar
Closeconnection_consultar
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label5_Click()

If MsgBox("Confirmar a exclusão do registro" & UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 0) & _
           " - " & UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 2) & _
           " - " & UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 3) & "?", vbYesNo + vbQuestion, "Deletar") = vbYes Then
    
    deletar.Connect_deletar
    deletar.Closeconnection_deletar
    
Else
    MsgBox (" A exclusão do registro foi cancelado com sucesso")
    
End If

Connect_consultar
Closeconnection_consultar



End Sub

Private Sub Label8_Click()
Unload UserForm2
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

UserForm3.Show

End Sub

Private Sub UserForm_Initialize()

cabecalho

UserForm2.ListBox1.ColumnWidths = "35; 50; 45; 60; 270; 70; 20"

With ComboBox1
    .AddItem "TODOS OS POSTOS"
    .AddItem "POSTO 1"
    .AddItem "POSTO 2"
    .AddItem "POSTO 3"
    .AddItem "POSTO 4"
End With

End Sub
