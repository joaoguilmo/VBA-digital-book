Attribute VB_Name = "consultar"
Sub Connect_consultar()


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

sql = "Select * From tabelageral"


If UserForm2.ComboBox1.Value <> "" Then


    sql = "select * from tabelageral where posto = " & "'" & UserForm2.ComboBox1.Value & "'" & ";"
    
End If

If UserForm2.ComboBox1.Value = "TODOS OS POSTOS" Then
    sql = "Select * From tabelageral"
End If




    
recordset.Open sql, Con, adOpenKeyset, adLockReadOnly
    
linhalistbox = 1
    


While Not recordset.EOF

   With UserForm2.ListBox1
               
            UserForm2.ListBox1.AddItem
            .List(linhalistbox, 0) = recordset(0).Value
            .List(linhalistbox, 1) = recordset(1).Value
            .List(linhalistbox, 2) = recordset(2).Value
            .List(linhalistbox, 3) = recordset(3).Value
            .List(linhalistbox, 4) = recordset(4).Value
            .List(linhalistbox, 5) = recordset(5).Value
            .List(linhalistbox, 6) = recordset(6).Value
            
   End With
   
   
   
        
        linhalistbox = linhalistbox + 1
        
        recordset.MoveNext
        

Wend

recordset.Close

    
End Sub

Sub Closeconnection_consultar()

    On Error Resume Next
    Con.Close
    Debug.Print "conexão fechada...."
    Set Con = Nothing
    Set Wbk = Nothing
    Set recordset = Nothing
    On Error GoTo 0

End Sub


