'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------


Public dbPath As String
Dim listProjects(50) As String
Dim listEmployees(50) As String
Dim listDesciptions(50) As String

'---------------------------------------------------------------------------------------
' Procedure : createReport
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub createReport()
    createReportSheet
    If checkDBConnection = False Then
        MsgBox " The DB Connection is not working sorry "
        findDB
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : createReportSheet
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :   checks if the Report Sheet does already
'               exist if so it gets deleted and then
'               a new sheet gets create named Report
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub createReportSheet()
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    
    For Each ws In Sheets
        If (ws.Name = "Report") Then
            ws.Delete
        End If
    Next
  
    
    Sheets.Add(Type:=xlWorksheet).Name = "Report"
    
    
   Application.DisplayAlerts = True
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : findDB
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub findDB()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkDBConnection
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkDBConnection() As Boolean
    checkDBConnection = False
    myPath = ActiveWorkbook.Path & "\Zeiterfassung" & "\pecoDB.accdb"
    
    If Dir(myPath) <> "" Then
         dbPath = myPath
        checkDBConnection = True
    Else
       
        checkDBConnection = False
    End If
    Debug.Print dbPath
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkTimesSheet
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :   Checks the Format of Sheet:Times Table1
'               Cells with wrong format :colored red
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkTimesSheet() As Boolean
    checkTimesSheet = True
    getProjects
    getMitarbeiterNames
    getTaetigkeitsarten
    Dim row As range
    
    ' function to check if input works
    
    For Each row In [Table1].Rows
        row.Cells.Interior.ColorIndex = 2   ' White = 1
      
        'Debug.Print row.Cells(1, 3).Value
        
        
        If checkDatum(row.Cells(1, 1).Value, row) = False Then
            row.Cells(1, 1).Interior.ColorIndex = 3  ' Red = 3
            checkTimesSheet = False
        End If
        
        If checkWochentag(row.Cells(1, 2).Value, row) = False Then
            row.Cells(1, 2).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkVon(row.Cells(1, 3).Text, row) = False Then
            row.Cells(1, 3).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkBis(row.Cells(1, 4).Text, row) = False Then
            row.Cells(1, 4).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkProjekt(row.Cells(1, 5).Value, row) = False Then
            row.Cells(1, 5).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkTaetigkeitsart(row.Cells(1, 6).Value, row) = False Then
            row.Cells(1, 6).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkMitarbeiter(row.Cells(1, 9).Value, row) = False Then
            row.Cells(1, 9).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkKW(row.Cells(1, 10).Value, row) = False Then
            row.Cells(1, 10).Interior.ColorIndex = 3
            checkTimesSheet = False
        End If
        
        If checkVon(row.Cells(1, 3).Text, row) = True And checkBis(row.Cells(1, 4).Text, row) = True And checkVonBis(row.Cells(1, 3).Text, row.Cells(1, 4).Text) = False Then
            row.Cells(1, 3).Interior.ColorIndex = 46
            row.Cells(1, 4).Interior.ColorIndex = 46
            checkTimesSheet = False
        End If
        
       
        If checkKW(row.Cells(1, 10).Value, row) = True And checkDatum(row.Cells(1, 1).Value, row) = True And checkKWCalculation(row.Cells(1, 1).Value, row.Cells(1, 10).Value) = False Then
            row.Cells(1, 10).Interior.ColorIndex = 46
            row.Cells(1, 1).Interior.ColorIndex = 46
            checkTimesSheet = False
        End If
        
        
    Next
    

End Function
'---------------------------------------------------------------------------------------
' Procedure : checkKWCalculation
' Author    : Anzumana
' Date      : 2/20/2014
' Purpose   :
' Inputs    :
' Returns   : Variant
'---------------------------------------------------------------------------------------

Function checkKWCalculation(enteredDate As Variant, kw As Variant)
    On Error GoTo checkKWCalculation_Error
    Dim i As Integer
    Debug.Print enteredDate
    Debug.Print kw
    
    Dim myRegExp As RegExp
    Dim myMatches As MatchCollection
    Dim myMatch As Match
    Set myRegExp = New RegExp
    Dim myDate As Date
    myRegExp.Pattern = "^(\d\d)\.(\d\d)\.(\d\d\d\d)$"
    Dim myString As String
   
    
    Set myMatches = myRegExp.Execute(enteredDate)
    
    For Each myMatch In myMatches
        
        
        myString = myMatch.SubMatches(1) + "/" + myMatch.SubMatches(0) + "/" + myMatch.SubMatches(2)
        Debug.Print myString
        myDate = CDate(myString)
        i = DatePart("ww", myDate)
                
                
        
    Next
   
    If (i <> kw) Then
        checkKWCalculation = False
    Else
        checkKWCalculation = True
    End If

   
    Exit Function

checkKWCalculation_Error:
    checkKWCalculation = False

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure checkKWCalculation of Module Module1"
    Exit Function
    
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkDatum
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkDatum(v As Variant, r As range) As Boolean
    Dim myRegExp As RegExp
  
   On Error GoTo checkDatum_Error

    Set myRegExp = New RegExp
   
    myRegExp.Pattern = "^(\d\d)\.(\d\d)\.(\d\d\d\d)$"
   

    If myRegExp.test(v) Then
        checkDatum = True
    Else
         checkDatum = False
    End If

   On Error GoTo 0
   Exit Function

checkDatum_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkDatum of Module Module1"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkWochentag
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkWochentag(v As Variant, r As range) As Boolean
   
    Dim myRegExp As RegExp
  
    Set myRegExp = New RegExp
   
    myRegExp.Pattern = "^Montag$|^Dienstag$|^Mittwoch$|^Donnerstag$|^Freitag$|^Samstag$|^Sonntag$"
   

    If myRegExp.test(v) Then
        checkWochentag = True
    Else
         checkWochentag = False
    End If
    
  
    
        
       
 
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkVon
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkVon(v As Variant, r As range) As Boolean
   Dim myRegExp As RegExp
   
   
  
    Set myRegExp = New RegExp
   
    myRegExp.Pattern = "(2[0-4]|((0|1)\d)):[0-5]\d"
   ' myRegExp.Pattern = "(2[0-4]|((0|1)*\d)):[0-5]\d"

    If myRegExp.test(v) Then
        checkVon = True
    Else
         checkVon = False
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : checkBis
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkBis(v As Variant, r As range) As Boolean
    Dim myRegExp As RegExp
  
    Set myRegExp = New RegExp
   
    myRegExp.Pattern = "(2[0-4]|((0|1)\d)):[0-5]\d"
    ' myRegExp.Pattern = "(2[0-4]|((0|1)*\d)):[0-5]\d"
   

    If myRegExp.test(v) Then
        checkBis = True
    Else
         checkBis = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkProjekt
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkProjekt(v As Variant, r As range) As Boolean
    Dim count As Integer
    count = 0
    While listProjects(count) <> ""
    If (listProjects(count) = v) Then
        checkProjekt = True
        Exit Function
    End If
    
       
        count = count + 1
    Wend
    
    

    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkTaetigkeitsart
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkTaetigkeitsart(v As Variant, r As range) As Boolean
    Dim count As Integer
    count = 0
    While listDesciptions(count) <> ""
    If (listDesciptions(count) = v) Then
        checkTaetigkeitsart = True
        Exit Function
    End If
    
        
        count = count + 1
    Wend
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkMitarbeiter
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkMitarbeiter(v As Variant, r As range) As Boolean
 Dim count As Integer
    count = 0
    While listEmployees(count) <> ""
    If (listEmployees(count) = v) Then
        checkMitarbeiter = True
        Exit Function
    End If
    
        
        count = count + 1
    Wend
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkKW
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   : Boolean
'---------------------------------------------------------------------------------------

Function checkKW(v As Variant, r As range) As Boolean
    
    
    Dim myRegExp As RegExp
    Set myRegExp = New RegExp
    myRegExp.Pattern = "(5[0-2])|([1-4][0-9])|(0[1-9])|[0-9]"
 

    If myRegExp.test(v) Then
        checkKW = True
    Else
         checkKW = False
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure :   checkVonBis
' Author    :   Anzumana
' Date      :   2/20/2014
' Purpose   :   logic timefrom has to be earlier then timeto. if conversion error function
'               return false also
' Inputs    :   timeto and timefrom
' Returns   :   Boolean
'---------------------------------------------------------------------------------------

Function checkVonBis(timefrom As Variant, timeto As Variant) As Boolean
    Dim c As Variant
    
    On Error GoTo checkVonBis_Error

    c = TimeValue(timeto) - TimeValue(timefrom)

    If c <= 0 Then
        checkVonBis = False
    Else
         checkVonBis = True
    End If

   
    Exit Function

checkVonBis_Error:
    checkVonBis = False

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure checkVonBis of Module Module1"
    Exit Function
    
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : CurrencyEx
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub CurrencyEx()
    Dim inputstr, re, amt
    Set re = New RegExp  'Create the RegExp object
    
    'Ask the user for the appropriate information
    inputstr = InputBox("I will help you convert USA and CAN currency. Please enter the amount to convert:")
    'Check to see if the input string is a valid one.
    re.Pattern = "(2[0-4]|((0|1)\d)):[0-5]\d"
    re.IgnoreCase = True
    Dim count As Integer
    count = 0
    Debug.Print re.test(inputstr)
    
    Do While re.test(inputstr) <> True
    'Prompt for another input if inputstr is not valid
    inputstr = InputBox("I will help you convert USA and GBP currency. Please enter the amount to(USD or GBP):")
    If count > 5 Then
        Exit Sub
    End If
    count = count + 1
    
    
    
    
    
    Loop

End Sub

'---------------------------------------------------------------------------------------
' Procedure : getProjects
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub getProjects()
    
     If Module1.checkDBConnection = True Then

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
    
    
    Rs1.Open "select * from Projekte where active = true", cn, adOpenKeyset, adLockOptimistic
    'projectName
    Dim count As Integer
    count = 0
    While Not Rs1.EOF
        Debug.Print Rs1.Fields(1).Value
        listProjects(count) = Rs1.Fields(1).Value
        count = count + 1
      
        Rs1.MoveNext
    
    Wend
    
    cn.Close
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getMitarbeiterNames
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub getMitarbeiterNames()
If Module1.checkDBConnection = True Then

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
    
    
    Rs1.Open "select * from Mitarbeiter", cn, adOpenKeyset, adLockOptimistic
    'projectName
    Dim count As Integer
    count = 0
    While Not Rs1.EOF
        Debug.Print Rs1.Fields(1).Value
        listEmployees(count) = Rs1.Fields(1).Value
        count = count + 1
      
        Rs1.MoveNext
    
    Wend
    
    cn.Close
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTaetigkeitsarten
' Author    : Anzumana
' Date      : 2/19/2014
' Purpose   :
' Inputs    :
' Returns   :
'---------------------------------------------------------------------------------------

Sub getTaetigkeitsarten()
    If Module1.checkDBConnection = True Then
    
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
        
        
        Rs1.Open "select * from WorkTypes", cn, adOpenKeyset, adLockOptimistic
        'projectName
        Dim count As Integer
        count = 0
        While Not Rs1.EOF
            Debug.Print Rs1.Fields(1).Value
            listDesciptions(count) = Rs1.Fields(1).Value
            count = count + 1
          
            Rs1.MoveNext
        
        Wend
        
        cn.Close
    End If
End Sub


