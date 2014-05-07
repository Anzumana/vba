
Private Sub CommandButton1_Click()
If checkMitarbeiterKeyPhrase(tb_keyPhrase.Value, lb_shorthands.Value) Then
    Sheet2.CommitToDatabase
    MsgBox "commit done"
    UserForm2.Hide
    
Else
MsgBox "Sorry but thats the wrong  Passphrase for the selected User"
End If

End Sub


Private Sub UserForm_Initialize()
    
    If (checkDBConnection) Then
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        sqlstring = "SELECT * FROM Mitarbeiter;"
        
        'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
    
        Set Rs1 = cn.Execute(sqlstring)
        While Not Rs1.EOF
        lb_shorthands.AddItem Rs1.Fields(1).Value
       
        Rs1.MoveNext
     
        Wend
        Rs1.Close
   
        cn.Close
    
    End If
End Sub

