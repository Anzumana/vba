Public xl0 As New Excel.Application
Public pptApp As Object
Public xlw As Object 'New Excel.Workbook
Public word As Object



Sub main()
' if i have an array of all the folders that i need i can iterate the function below over that array to get what we wanted.
        editPropertiesOfFolder (ActiveWorkbook.Path)
End Sub
Sub editPropertiesOfFolder(folderPath As String)


    MyDir = folderPath
    strPath = MyDir & ":" 'on mac :
    Debug.Print strPath
    strFile = Dir(strPath)
    'Loop through each file in the folder
    
    Do While Len(strFile) > 0
        ' if file found is active workbook then next
        
        If Right(strFile, 4) = "xlsx" Then 'xlsx
        Debug.Print strPath & strFile
            Set xlw = xl0.Workbooks.Open(strPath & strFile)
             fillGui (strFile)
        ElseIf Right(strFile, 4) = "docx" Then 'docx
            On Error Resume Next
            Set word = GetObject(, "word.application") 'gives error 429 if Word is not open
            If Err = 429 Then
                Set doc = CreateObject("word.application") 'creates a Word application
                Err.Clear
            End If
            word.Visible = False
            Set doc = word.documents.Open(strPath & strFile)
             fillGui (strFile)
        ElseIf Right(strFile, 4) = "pptx" Then 'pptx
            Set pptApp = CreateObject("PowerPoint.Application")
            Set ppt = pptApp.Presentations.Open(strPath & strFile)
             fillGui (strFile)
        End If

        strFile = Dir
    Loop
       
End Sub
Sub fillGui(filename As Variant)


            frmDisplay.lblFilename = filename
            frmDisplay.txtTitle = xlw.BuiltinDocumentProperties.Item("Title")
            frmDisplay.txtSubject = xlw.BuiltinDocumentProperties.Item("Subject")
            frmDisplay.txtAuthor = xlw.BuiltinDocumentProperties.Item("Author")
            frmDisplay.txtManager = xlw.BuiltinDocumentProperties.Item("Manager")
            frmDisplay.txtCompany = xlw.BuiltinDocumentProperties.Item("Company")
            frmDisplay.txtCategory = xlw.BuiltinDocumentProperties.Item("Category")
            frmDisplay.txtKeywords = xlw.BuiltinDocumentProperties.Item("Keywords")
            frmDisplay.txtComments = xlw.BuiltinDocumentProperties.Item("Comments")
            frmDisplay.txtHyperlinkBase = xlw.BuiltinDocumentProperties.Item("Hyperlink Base")
            frmDisplay.Show
            'xlw.Save
            'xlw.Close
            Set xl0 = Nothing
            Set xlw = Nothing
End Sub
