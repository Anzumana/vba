'---------------------------------------------------------------------------------------
' Procedure : Update_click
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub Update_click()
    UserForm1.cmd2_Click
    
End Sub
  
'---------------------------------------------------------------------------------------
' Procedure : test
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub test()
    Debug.Print myPath
    
    'Module1.checkTable
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CommitToDatabase
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------
  Sub CommitToDatabase()
    versionNumber = "1.1"
    Module1.checkVersionNumber (versionNumber)
    
    If Module1.checkDBConnection = True And Module1.checkTimesSheet = True Then 'check if date of commited values are ok
    
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
  
    'sqlstring = "INSERT INTO Table1(myNumber)VALUES(4)"
    'sqlstring = "INSERT INTO Zeiterfassung Values = Table1"
    'sql2string = "SELECT * FROM Mitarbeiter WHERE (((Mitarbeiter.[shorthand])=""CD""));"
    
    
    Rs1.Open "select * from Zeiterfassung", cn, adOpenKeyset, adLockOptimistic
    
    Dim fieldsArray(8) As Variant
    Dim values(8) As Variant
    
    fieldsArray(0) = "Datum"
    fieldsArray(1) = "Wochentag"
    fieldsArray(2) = "Von"
    fieldsArray(3) = "Bis"
    fieldsArray(4) = "Projekt"
    fieldsArray(5) = "T?tigkeitsart"
    fieldsArray(6) = "T?tigkeitsbeschreibung"
    fieldsArray(7) = "Mitarbeiter"
    fieldsArray(8) = "KW"
    
    values(0) = #1/1/1900#
    values(1) = "Mittwoch"
    values(2) = 1 / 10 / 2014
    values(3) = 1 / 1 / 2011
    values(4) = "Dienstag"
    values(5) = "Dienstag"
    values(6) = "Dienstag"
    values(7) = "Dienstag"
    values(8) = "Dienstag"
    
    
    Dim row As range
    Dim myRegExp As RegExp
    Dim myMatches As MatchCollection
    Dim myMatch As Match
    Set myRegExp = New RegExp
    Dim myDate As Date
    myRegExp.Pattern = "(\d\d)\.(\d\d)\.(\d\d\d\d)"
    Dim myString As String
    
    ' function to check if input works
    
    For Each row In [Table1].Rows
        'Debug.Print row.Columns(1).Value
        'Debug.Print myRegExp.test(row.Columns(1).Value)
        Set myMatches = myRegExp.Execute(row.Columns(1).Value)
       
    
    For Each myMatch In myMatches
        
        'Debug.Print "#" + myMatch.SubMatches(1) + "/" + myMatch.SubMatches(0) + "/" + myMatch.SubMatches(2) + "#"
        'If myRegExp.test(row.Columns(1).Value) Then
        myString = myMatch.SubMatches(1) + "/" + myMatch.SubMatches(0) + "/" + myMatch.SubMatches(2)
        Debug.Print myString
        myDate = CDate(myString)
        values(0) = myDate
                
                
                
        'Debug.Print myDate
        'values(0) = myDate
        'End If
        Next
        'Debug.Print myMatch.SubMatches(1)
        'values(0) = row.Columns(1).Value 'Datum
        'Debug.Print Format(row.Columns(1).Value)
        'Debug.Print Format(37973, "yyyy/mm/dd")
        'Debug.Print DateValue(row.Columns(1).Value)
       
        'Debug.Print TimeValue(row.Columns(1).Value)
        values(1) = row.Columns(2).Value     ' Wochentag
        values(2) = row.Columns(3).Value     ' Von
        values(3) = row.Columns(4).Value     ' Bis
        values(4) = row.Columns(5).Value     ' Projekt
        values(5) = row.Columns(6).Value     ' Taetigkeitsart
        values(6) = row.Columns(7).Value     'Taetigkeitsbeschreibung
        values(7) = row.Columns(9).Value     ' Mitarbeiter
        values(8) = row.Columns(10).Value    'KW
        Rs1.AddNew fieldsArray, values
        Rs1.Update
    Next
    'Dim myDate As Date
    'myDate = #12/23/2013#
    'Debug.Print myDate
    'Debug.Print IsDate(myDate)
    'Dim LValue As String

    'LValue = Format(myDate, "mm/dd/yyyy")
    'Debug.Print LValue
    
   
    
        cn.Close
    End If
  End Sub


