
' Zeiterfassung Verion 1.1  last Update : 28.11.2013
Sub getTimes()
' This sub adds the default times to for our time combo boxes
    With cb_timeto
        For x = 0 To 24
        .AddItem x 'CStr(x) & ":" & "00"
        Next x
    End With
    With cb_timefrom
        For x = 0 To 24
        .AddItem x 'CStr(x) & ":" & "00"
        Next x
    End With
      With cb_timefrommin
        .AddItem "00"
        .AddItem "15"
        .AddItem "30"
        .AddItem "45"
        End With
    With cb_timetomin
        .AddItem "00"
        .AddItem "15"
        .AddItem "30"
        .AddItem "45"
        End With
End Sub

Sub getWorkersNames()
        MyPath = ActiveWorkbook.Path & "\Zeiterfassung" & "\pecoDB.accdb"
    
    If Dir(MyPath) <> "" Then
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        sqlstring = "SELECT * FROM Mitarbeiter;"
        
        'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & MyPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
    End If
    
    Set Rs1 = cn.Execute(sqlstring)
 
 
 
   
    While Not Rs1.EOF
        lb_workers.AddItem Rs1.Fields(1).Value
      
        Rs1.MoveNext
     
   Wend
   Rs1.Close
   
    cn.Close
   

    


End Sub

Sub getProjectNames()
        MyPath = ActiveWorkbook.Path & "\Zeiterfassung" & "\pecoDB.accdb"
    
    If Dir(MyPath) <> "" Then
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        sqlstring = "SELECT * FROM Projekte  where active = true;"
        
        'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & MyPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
    End If
    
    Set Rs1 = cn.Execute(sqlstring)
   
   While Not Rs1.EOF
        lb_projects.AddItem Rs1.Fields(1).Value
      Rs1.MoveNext
    
   Wend

    cn.Close
    
    
End Sub
Public Sub cmd2_Click()
' the fertig button on the user form
' update times on the times sheet
    Application.ThisWorkbook.Worksheets("Times").Activate
    Dim weeknumber As String
    Dim weeknumbertmp As String
    
    Dim counter As Integer
    Dim countermax As Integer
    Dim weektotal As Date
    Dim weektotalf As Date
    
    
    counter = 1
    countermax = 0
  
    
    While Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 10), Cells(counter, 10)).Value <> ""
        countermax = countermax + 1
        counter = counter + 1
    Wend
    
    If countermax = 1 Then
        UserForm1.Hide
        Exit Sub
    End If
    
    
    'clear calculated values
        
    counter = 2
    While counter <= countermax
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 11), Cells(counter, 11)).Value = ""
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 12), Cells(counter, 12)).Value = ""
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 13), Cells(counter, 13)).Value = ""
        counter = counter + 1
    Wend
    

    ' end clear
    
    
      counter = 2
Dim daytotal As Date
    daytotal = 0
  ' calculate hourse per not per day per entry
  While counter <= countermax
        
        
         daytotal = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 4), Cells(counter, 4)).Value - Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 3), Cells(counter, 3)).Value
            
            
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).NumberFormat = "[h]:mm:ss"
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value = daytotal
            
        counter = counter + 1
    Wend
    
    
    counter = 2
    'get first weeknumber
    
    
    weeknumber = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 10), Cells(counter, 10)).Value
    
    weektotal = 0
    weektotalf = 0
    While counter <= countermax
        weeknumbertmp = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 10), Cells(counter, 10)).Value
        If weeknumber <> weeknumbertmp Then
            
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 12), Cells(counter - 1, 12)) = weektotal
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 12), Cells(counter - 1, 12)).NumberFormat = "[h]:mm:ss"
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 13), Cells(counter - 1, 13)) = weektotalf
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 13), Cells(counter - 1, 13)).NumberFormat = "[h]:mm:ss"
            
            
            weektotal = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value
            If Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 5), Cells(counter, 5)).Value <> "Intern" Then
                weektotalf = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value
                Else
                weektotalf = 0
                
            End If
            
            weeknumber = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 10), Cells(counter, 10)).Value
        
        Else
                weektotal = weektotal + Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value
               
                If Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 5), Cells(counter, 5)).Value <> "Intern" Then
                     weektotalf = weektotalf + Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value
                End If
                
                
                
        End If
        
        counter = counter + 1
    Wend
   
        
       
    
    weeknumber = Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermax, 10), Cells(countermax, 10)).Value
    
   
    weektotal = 0
    weektotalf = 0
    countermaxold = countermax

    While countermax >= 0
    
        weeknumbertmp = Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermax, 10), Cells(countermax, 10)).Value
        
        
        
        
        If weeknumber <> weeknumbertmp Then
         
            
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermaxold, 12), Cells(countermaxold, 12)) = weektotal
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermaxold, 12), Cells(countermaxold, 12)).NumberFormat = "[h]:mm:ss"
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermaxold, 13), Cells(countermaxold, 13)) = weektotalf
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermaxold, 13), Cells(countermaxold, 13)).NumberFormat = "[h]:mm:ss"
          countermax = 0
        
        Else
           
            
            weektotal = weektotal + Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermax, 8), Cells(countermax, 8)).Value
            If Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermax, 5), Cells(countermax, 5)).Value <> "Intern" Then
                    weektotalf = weektotalf + Application.ThisWorkbook.Worksheets("Times").Range(Cells(countermax, 8), Cells(countermax, 8)).Value
            End If
            
            
        End If
        
        countermax = countermax - 1
       
    
        
    Wend
    countermax = countermaxold
    counter = 0
    Dim datestring As String
    Dim datestringtmp As String
    
    
    
    
    counter = 2
    datestring = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 1), Cells(counter, 1)).Value
    daytotal = 0
    While counter <= countermax
        datestringtmp = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 1), Cells(counter, 1)).Value
        
        If datestringtmp = datestring Then
            daytotal = daytotal + Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value
            
        Else
            
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 11), Cells(counter - 1, 11)).NumberFormat = "[h]:mm:ss"
            Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 11), Cells(counter - 1, 11)).Value = daytotal
            
            daytotal = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 8), Cells(counter, 8)).Value
            datestring = Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 1), Cells(counter, 1)).Value
        End If
        
        counter = counter + 1
    Wend
    
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 11), Cells(counter - 1, 11)).NumberFormat = "[h]:mm:ss"
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter - 1, 11), Cells(counter - 1, 11)).Value = daytotal
    
    
  
    
    UserForm1.Hide
    
    
End Sub
Sub cmd1_click()
    Dim a As Boolean
    a = checkuserform
    ' check values inside of the user form
    
    If a = True Then
        CommandButton1_Click
        MsgBox " Daten wurden hinzugefŸgt"
    Else
        
    End If
      
End Sub
Function checkuserform() As Boolean
    checkuserform = True
    
    
  
    Debug.Print lb_type.Value, lb_projects; lb_workers
    If lb_type.Value = "" Then
        MsgBox "Es wurde keine T_tigkeitsart angegeben"
        checkuserform = False
        Exit Function
  
    End If
    
    If IsNull(lb_projects) Then
        MsgBox "Es wurde kein Projekt ausgew_hlt"
        checkuserform = False
        Exit Function
    End If
    
    
    If IsNull(lb_workers) Then
        MsgBox " Es wurde keine Mitarbeiter ausgew_hlt "
        checkuserform = False
        Exit Function
        
    End If
      
    
    If tb_description = "" Then
        MsgBox "Es wurde keine Beschriebung angegeben"
        checkuserform = False
        Exit Function
    End If
     If tb_date = "" Then
        MsgBox " Es wurde keine Datum angeben"
        checkuserform = False
        Exit Function
    End If
 
    If cb_timefrom = "" Then
        MsgBox "Es wurde keine Anfangszeit angegeben"
        checkuserform = False
        Exit Function
    End If
       If cb_timefrommin = "" Then
        MsgBox "Es wurde keine Anfangszeit angegeben"
        checkuserform = False
        Exit Function
    End If
    
    If cb_timeto = "" Then
        MsgBox "Es wurde keine Endzeit angegeben"
        checkuserform = False
        Exit Function
    End If
      If cb_timetomin = "" Then
        MsgBox "Es wurde keine Endzeit angegeben"
        checkuserform = False
        Exit Function
    End If
    
    Dim a As String
    Dim b As String
    a = cb_timefrom.Value & ":" & cb_timefrommin.Value
    'a = cb_timefrom.Value
    b = cb_timeto.Value & ":" & cb_timetomin.Value

   ' b = cb_timeto.Value
    Dim c As Variant

    c = TimeValue(b) - TimeValue(a)
    If c <= 0 Then
        MsgBox "Die Anfangszeit muss vor der Endzeit liegen."
        checkuserform = False
        Exit Function
    End If
    If Not IsDate(tb_date.Value) Then
        MsgBox "†berprŸfe das Datumformat"
        checkuserform = False
        Exit Function
    End If
    
End Function
 Sub CommandButton1_Click()
    ' button hinzufuegen
    
    ' check if the values entered by the user are correct
   
    
    
    
    
    
    
    ' put everything onto sheet times that was entered into the userform
    lastrow
    
    Dim counter As Integer
    counter = 1
    weekdaycalc
    Select Case weekdaycalc
        Case "Montag"
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 1), Cells(counter + lastrow, 1)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 2), Cells(counter + lastrow, 2)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 3), Cells(counter + lastrow, 3)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 4), Cells(counter + lastrow, 4)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 5), Cells(counter + lastrow, 5)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 6), Cells(counter + lastrow, 6)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 7), Cells(counter + lastrow, 7)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 8), Cells(counter + lastrow, 8)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 9), Cells(counter + lastrow, 9)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 10), Cells(counter + lastrow, 10)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 10), Cells(counter + lastrow, 11)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 10), Cells(counter + lastrow, 12)).Interior.ColorIndex = 19
        Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 10), Cells(counter + lastrow, 13)).Interior.ColorIndex = 19
   
    End Select
    
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 10), Cells(counter + lastrow, 10)).Value = DatePart("ww", tb_date.Value)
    Dim a As String
    Dim b As String
   ' a = cb_timefrom.Value
    'b = cb_timeto.Value
    Dim c As Variant




    a = cb_timefrom.Value & ":" & cb_timefrommin.Value
 
    b = cb_timeto.Value & ":" & cb_timetomin.Value





    c = TimeValue(b) - TimeValue(a)
     Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 8), Cells(counter + lastrow, 8)).NumberFormat = "[h]:mm:ss"
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 8), Cells(counter + lastrow, 8)).Value = c
    
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 6), Cells(counter + lastrow, 6)).Value = lb_type.Value
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 7), Cells(counter + lastrow, 7)).Value = tb_description.Value
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 5), Cells(counter + lastrow, 5)).Value = lb_projects.Value
     Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 9), Cells(counter + lastrow, 9)).Value = lb_workers.Value
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 3), Cells(counter + lastrow, 3)).Value = a
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 4), Cells(counter + lastrow, 4)).Value = b
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 2), Cells(counter + lastrow, 2)).Value = weekdaycalc
    Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter + lastrow, 1), Cells(counter + lastrow, 1)).Value = tb_date.Value

End Sub

Function weekdaycalc() As String
    Dim a As Integer
    
   a = weekday(tb_date, 2)
 
   
    Select Case a
            Case 1
        weekdaycalc = "Montag"
        Case 2
        weekdaycalc = "Dienstag"
        Case 3
        weekdaycalc = "Mittwoch"
        Case 4
        weekdaycalc = "Donnerstag"
        Case 5
        weekdaycalc = "Freitag"
        Case 6
        weekdaycalc = "Samstag"
        Case 7
        weekdaycalc = "Sonntag"
        
    End Select
   
    
End Function

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    checkDBConnection
    getProjectNames
    getWorkersNames
    getTimes
    getTypes
    getDate
End Sub
Sub checkDBConnection()
Dim MyPath As String

    MyPath = ActiveWorkbook.Path & "\Zeiterfassung" & "\pecoDB.accdb"
    If Dir(MyPath) <> "" Then
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        sqlstring = "SELECT * FROM WorkTypes;"
        
        'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & MyPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
        cn.Close
    Else
        MsgBox " Leider konnte die Datenbank nicht gefunden werdeen "
    End If
    
End Sub
Sub getDate()
    tb_date.Value = Date
    
End Sub
Sub getTypes()
      MyPath = ActiveWorkbook.Path & "\Zeiterfassung" & "\pecoDB.accdb"
   
    If Dir(MyPath) <> "" Then
        Dim cmd As New ADODB.Command
        Dim AccessConnect As String
        Dim Rs1 As New Recordset
        Dim sqlstring As String
  
        Set cn = New ADODB.Connection
        sqlstring = "SELECT * FROM WorkTypes;"
        
        'AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Anzumana\Dropbox\peco\pecoDB.accdb;Jet OLEDB:Database Password=test;"
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & MyPath & ";Jet OLEDB:Database Password=test;"
        ' Note! Reports say that a database encrypted using Access 2010 - 2013 default encryption scheme does not work with this connection string.
        ' In Access; try options and choose 2007 encryption method instead.
        ' That should make it work. We do not know of any other solution. Please get in touch if other solutions is available!
        cn.ConnectionString = AccessConnect
    
        cn.Open
    End If
    
    Set Rs1 = cn.Execute(sqlstring)
   
   While Not Rs1.EOF
        lb_type.AddItem Rs1.Fields(1).Value
      Rs1.MoveNext
    
   Wend
    
    cn.Close
    
End Sub
Function lastrow() As Integer
    Dim counter, countermax As Integer
    counter = 1
    countermax = 0
    Application.ThisWorkbook.Worksheets("Times").Activate
    
    While Application.ThisWorkbook.Worksheets("Times").Range(Cells(counter, 1), Cells(counter, 1)).Value <> ""
        countermax = countermax + 1
        counter = counter + 1
    Wend
    lastrow = countermax
End Function

