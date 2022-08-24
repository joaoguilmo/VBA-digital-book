Attribute VB_Name = "Salvar"
Sub Connect()

Dim Con As ADODB.Connection
Dim Wbk As Workbook
Dim Db  As String
Dim recordset As New ADODB.recordset

Set Wbk = PastaWb

Db = Wbk.Path & "\db.accdb"
Set Con = New ADODB.Connection

    If Con.State = 0 Then

        
        Con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & Db & ";" & _
                                "Jet OLEDB:Database Password=MyDbPassword;"
        
        Con.Open
        Debug.Print "conectado...."

    End If

Dim sql As String

sql = "Select * From tabelageral"
    
recordset.Open sql, Con, adOpenKeyset, adLockOptimistic
    
recordset.AddNew
    
recordset.Fields("timestamp").Value = Now
recordset.Fields("posto").Value = UserForm1.ComboBox1.Value
recordset.Fields("equipamento").Value = UserForm1.TextBox1.Value
recordset.Fields("descricao").Value = UserForm1.TextBox2.Value
recordset.Fields("notamanutencao").Value = UserForm1.TextBox3.Value
recordset.Fields("concertado").Value = False
recordset.Fields("username").Value = Application.UserName
    
recordset.Update
recordset.Close
    
Debug.Print "Dados gravados"


    
End Sub

Sub Closeconnection()

    On Error Resume Next
    Con.Close
    Debug.Print "conexão fechada...."
    Set Con = Nothing
    Set Wbk = Nothing
    Set recordset = Nothing
    
    On Error GoTo 0

End Sub

