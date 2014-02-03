Sub test()
  Dim cn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Errs1 As Errors
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim AccessConnect As String
    Dim Rs1 As New Recordset
  
    Dim sqlstring As String
    'sqlstring = "INSERT INTO Table1(myNumber)VALUES(4)"
    sqlstring = "INSERT INTO Zeiterfassung (Wochentag)VALUES(""Testrightnow"")"
    'sql2string = "SELECT * FROM Mitarbeiter WHERE (((Mitarbeiter.[shorthand])=""CD""));"
    sql2string = "SELECT * FROM Mitarbeiter;"
    Set cn = New ADODB.Connection
    
  
    
    'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Persist Security Info=False;"
    AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
    ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
    ' In Access; try options and choose 2007 encryption method instead.
    ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
    cn.ConnectionString = AccessConnect
    
    cn.Open
    'cn.Execute sqlstring ' this is for putting sql statements in
  
    Set Rs1 = cn.Execute(sql2string) ' creating the record set to put content onto worksheets
    'Application.ThisWorkbook.Worksheets("Sheet1").Range(Cells(1, 7), Cells(1, 7)).Value = Rs1.Fields(2).Value
 
  Dim RowCnt, FieldCnt As Integer
   RowCnt = 1
   ' Use field names as headers in the first row.
   For FieldCnt = 0 To Rs1.Fields.Count - 1
      Cells(RowCnt, FieldCnt + 1).Value = Rs1.Fields(FieldCnt).Name
      Rows(1).Font.Bold = True
   Next FieldCnt
   
    'Fill rows with records, starting at row 2.
   RowCnt = 2
   
   While Not Rs1.EOF
      For FieldCnt = 0 To Rs1.Fields.Count - 1
         Cells(RowCnt, FieldCnt + 1).Value = _
         Rs1.Fields(FieldCnt).Value
      Next FieldCnt
      Rs1.MoveNext
      RowCnt = RowCnt + 1
   Wend
   
   
    
  
    cn.Close
   

    
End Sub

