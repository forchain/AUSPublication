Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub ImportAuthor()

    Dim y As Integer
    Dim i As Integer
    Dim sPath As String
    Dim sSheet As String

    ' Faculty In
    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_FACULTY_IN_FILE)
    sSheet = Config.Val(Consts.SECTION_AUTHOR, Consts.KEY_FACULTY_IN_SHEET) & "!"

    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "LinkFacultyIn", sPath, True, sSheet
    
    ' Faculty Out
    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_FACULTY_OUT_FILE)
    sSheet = Config.Val(Consts.SECTION_AUTHOR, Consts.KEY_FACULTY_OUT_SHEET) & "!"

    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "LinkFacultyOut", sPath, True, sSheet
    
    ' Senior
    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_SENIOR_FILE)
    sSheet = Config.Val(Consts.SECTION_AUTHOR, Consts.KEY_SENIOR_SHEET) & "!"

    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "LinkSenior", sPath, True, sSheet
        
    '  Staff
    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_STAFF_FILE)
    sSheet = Config.Val(Consts.SECTION_AUTHOR, Consts.KEY_STAFF_SHEET) & "!"

    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "LinkStaff", sPath, True, sSheet
    
    ' College
    If Not CheckTable("College") Then
        CurrentDb.Execute "CreateCollege", dbFailOnError
        Debug.Print "CreateCollege", CurrentDb.RecordsAffected
    
        CurrentDb.Execute "InsertUnknownCollege", dbFailOnError
        Debug.Print "InsertUnknownCollege", CurrentDb.RecordsAffected
    End If

    CurrentDb.Execute "InsertCollege", dbFailOnError
    Debug.Print "InsertCollege", CurrentDb.RecordsAffected
    
    ' Department
    If Not CheckTable("Department") Then
        CurrentDb.Execute "CreateDepartment", dbFailOnError
        Debug.Print "CreateDepartment", CurrentDb.RecordsAffected
    
        CurrentDb.Execute "InsertUnknownDepartment", dbFailOnError
        Debug.Print "InsertUnknownDepartment", CurrentDb.RecordsAffected
    End If
    CurrentDb.Execute "InsertDepartment", dbFailOnError
    Debug.Print "InsertDepartment", CurrentDb.RecordsAffected
    
    ' Job
    If Not CheckTable("Job") Then
        CurrentDb.Execute "CreateJob", dbFailOnError
        Debug.Print "CreateJob", CurrentDb.RecordsAffected
    
        CurrentDb.Execute "InsertUnknownJob", dbFailOnError
        Debug.Print "InsertUnknownJob", CurrentDb.RecordsAffected
    End If
    CurrentDb.Execute "InsertJob", dbFailOnError
    Debug.Print "InsertJob", CurrentDb.RecordsAffected
    
    ' Author
    
    If Not CheckTable("Author") Then
        CurrentDb.Execute "CreateAuthor", dbFailOnError
        Debug.Print "CreateAuthor", CurrentDb.RecordsAffected
        
        CurrentDb.Execute "IndexAuthor", dbFailOnError
        Debug.Print "IndexAuthor", CurrentDb.RecordsAffected
    End If
    
    
    CurrentDb.Execute "InsertAuthor", dbFailOnError
    Debug.Print "InsertAuthor", CurrentDb.RecordsAffected
    
    DoCmd.DeleteObject acTable, "LinkFacultyIn"
    Debug.Print "Delete LinkFacultyIn", CurrentDb.RecordsAffected
    DoCmd.DeleteObject acTable, "LinkFacultyOut"
    Debug.Print "Delete LinkFacultyOut", CurrentDb.RecordsAffected
    DoCmd.DeleteObject acTable, "LinkSenior"
    Debug.Print "Delete LinkSenior", CurrentDb.RecordsAffected
    DoCmd.DeleteObject acTable, "LinkStaff"
    Debug.Print "Delete LinkStaff", CurrentDb.RecordsAffected
   
    'DoCmd.OpenTable "Author"
    'Debug.Print "TestImportAuthor"
End Sub

Public Sub ImportPaper()

    Dim currYear As Integer
    currYear = Year(Date)

    Dim y As Integer
    Dim i As Integer
    Dim sKey As String
    Dim sPath As String
    
    If Not CheckTable("Paper") Then
        CurrentDb.Execute "CreatePaper", dbFailOnError
        Debug.Print "CreatePaper", CurrentDb.RecordsAffected
    End If


    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.QueryDefs("InsertPaper")
    
    For y = Consts.BEIGN_YEAR To currYear
        For i = 0 To UBound(Consts.INDICES)
            
            sKey = Config.IndexKey(Consts.INDICES(i), y)

            sPath = Config.SheetPath(Consts.SECTION_INDEX, sKey)
            
            'Debug.Print path

            DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, "LinkPaper", sPath, True, Consts.SHEET_PAPER & "!"
            
            qd.Parameters("Year").Value = y
            qd.Parameters("Index").Value = i + 1

            qd.Execute dbFailOnError
            Debug.Print "InsertPaper", CurrentDb.RecordsAffected
        
            DoCmd.DeleteObject acTable, "LinkPaper"
            Debug.Print "Delete LinkPaper", CurrentDb.RecordsAffected
        Next i
    Next y
    ' UnknownPaper
'
'    sPath = Config.SheetPath(Consts.SECTION_PAPER, Consts.KEY_UNKNOWN_PAPER_FILE)
'
'    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownPaper", sPath, True, Consts.SHEET_UNKNOWN_PAPER
'
'
'    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "UnknownPaper", sPath, True, Consts.SHEET_UNKNOWN_PAPER & "!"
'
End Sub

Public Sub FillWeight()
    
    Dim db     As DAO.Database
    Set db = CurrentDb
  
    db.Execute "CreateWeight", dbFailOnError
    Debug.Print "CreateWeight", db.RecordsAffected
    

    Dim rsPaper As Recordset
    Set rsPaper = db.OpenRecordset("Paper", dbOpenTable)
'    Set rsPaper = db.OpenRecordset("ViewPaper", dbOpenDynaset)

    Dim qd As DAO.QueryDef
    Dim an As String
    Set qd = db.QueryDefs("InsertWeight")

    Do While Not rsPaper.EOF
        Dim authors() As String
        '        If rsPaper!Id = 828 Then
        '            Debug.Print rsPaper!AuthorNames
        '        End If
       
        If IsNull(rsPaper!AuthorNames) Then
            qd.Parameters("PaperID").Value = rsPaper!Id
            qd.Parameters("AuthorName").Value = ""
            qd.Execute dbFailOnError
        Else
            authors = Split(rsPaper!AuthorNames, ";")
            Dim iI As Integer
            For iI = 0 To UBound(authors)
                an = authors(iI)
                qd.Parameters("PaperID").Value = rsPaper!Id
                qd.Parameters("AuthorName").Value = Paper.FixName(an)
                qd.Execute dbFailOnError
            Next iI
        End If


        rsPaper.MoveNext
    Loop
    
    ' Unknown Author
'    Dim sPath As String
'    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_UNKNOWN_AUTHOR_FILE)
'
'    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownAuthor", sPath, True, Consts.SHEET_UNKNOWN_AUTHOR
End Sub
Public Function ImportSheets()
    App.ClearTables

    ImportPaper
    
    ImportAuthor
    
    FillWeight
    
    Dim rResult As VbMsgBoxResult
    rResult = MsgBox("Done", vbYes, "Import Sheets")
End Function


