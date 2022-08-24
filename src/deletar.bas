Attribute VB_Name = "deletar"
Sub Connect_deletar()


Dim Con As ADODB.Connection
Dim Wbk As Workbook
Dim Db  As String
Dim recordset As New ADODB.recordset
Dim linhalistbox As Integer

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

    
If UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 0) <> "" Then

    sql = "select * from tabelageral where Codigo = " & UserForm2.ListBox1.List(UserForm2.ListBox1.ListIndex, 0) & ";"
    
End If


    
recordset.Open sql, Con, adOpenDynamic, adLockOptimistic

recordset.Delete

recordset.Close


    
End Sub

Sub Closeconnection_deletar()

    On Error Resume Next
    Con.Close
    Debug.Print "conexão fechada...."
    Set Con = Nothing
    Set Wbk = Nothing
    Set recordset = Nothing
    On Error GoTo 0

End Sub



