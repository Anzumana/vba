  Sub Update_click()
    UserForm1.cmd2_Click
    
  End Sub
  
  
  Sub CommitToDatabase()
     MyPath = ActiveWorkbook.Path & "\Zeiterfassung" & "\pecoDB.accdb"
    If Dir(MyPath) <> "" Then
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        
        
        'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & MyPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
  End If
    'sqlstring = "INSERT INTO Table1(myNumber)VALUES(4)"
    sqlstring = "INSERT INTO Zeiterfassung Values = Table1 "
    'sql2string = "SELECT * FROM Mitarbeiter WHERE (((Mitarbeiter.[shorthand])=""CD""));"
    
    
    Rs1.Open "select * from Zeiterfassung", cn, adOpenKeyset, adLockOptimistic
    Dim fieldsArray(8) As Variant
    fieldsArray(0) = "Datum"
    fieldsArray(1) = "Wochentag"
    fieldsArray(2) = "Von"
    fieldsArray(3) = "Bis"
    fieldsArray(4) = "Projekt"
    fieldsArray(5) = "Tätigkeitsart"
    fieldsArray(6) = "Tätigkeitsbeschreibung"
    fieldsArray(7) = "Mitarbeiter"
    fieldsArray(8) = "KW"
    Dim values(8) As Variant
    values(0) = 1 / 10 / 2014
    values(1) = "Mittwoch"
    values(2) = 1 / 10 / 2014
    values(3) = 1 / 1 / 2011
    values(4) = "Dienstag"
    values(5) = "Dienstag"
    values(6) = "Dienstag"
    values(7) = "Dienstag"
    values(8) = "Dienstag"
    
    Dim row As Range
    
    For Each row In [Table1].Rows
        values(0) = row.Columns(1).Value 'Datum
       Debug.Print Format(row.Columns(1).Value)
        'Debug.Print Format(37973, "yyyy-mm-dd")
       
       
       'Debug.Print TimeValue(row.Columns(1).Value)
       values(1) = row.Columns(2).Value     ' Wochentag
       values(2) = row.Columns(3).Value     ' Von
       values(3) = row.Columns(4).Value     ' Bis
       values(4) = row.Columns(5).Value     ' Projekt
       values(5) = row.Columns(6).Value     ' Taetigkeitsart
       values(6) = row.Columns(7).Value     'Taetigkeitsbeschreibung
       values(7) = row.Columns(9).Value     ' Mitarbeiter
       values(8) = row.Columns(10).Value    'KW
       ' Rs1.AddNew fieldsArray, values
        'Rs1.Update
    Next
    'Dim myDate As Date
    'myDate = #12/23/2013#
    'Debug.Print myDate
    'Debug.Print IsDate(myDate)
    'Dim LValue As String

    'LValue = Format(myDate, "mm/dd/yyyy")
    'Debug.Print LValue
  
    cn.Close
  End Sub

