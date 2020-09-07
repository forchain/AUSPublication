Option Compare Database
Option Explicit


Sub DeleteMFileV3()

    Dim fso    As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject
    
    Dim p As String
    p = CurrentProject.path
    
    fso.MoveFile p & ".\unimported.txt", p & ".\imported.txt"
    Debug.Print "Renamed to imported.txt"

End Sub

Sub ExportPaperErrorV3()
    Dim filepath As String
    filepath = CurrentProject.path & ".\PaperError.xlsx"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "ErrorPaperISSN", filepath, True, ""
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "ErrorPaperYear", filepath, True, ""
End Sub

Public Function DebugDB()
    ImportSheetV3
    'ImportPaperV3
    'ImportAuthorV3
End Function

Public Function TestImportPaperV3()
    CloseAllTablesV3
    DeleteAllRelationsV3
    DeleteAllTablesV3
    
    ImportJournalV3
    
    Set fso = New Scripting.FileSystemObject
    Dim paperFile As String
    paperFile = CurrentProject.path & "\paper.xlsx"
    
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "RawPaper", paperFile, True, "2018!"
    'DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "RawPaper", paperFile, True, "2019!"

    'DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Sheet1", paperFile, True, "Sheet1!A:B"
    'DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Sheet1", paperFile, True, "Sheet1!C:D"
End Function

Public Sub TestUnknownJournal()
    CloseAllTablesV3

    Set fso = New Scripting.FileSystemObject
    Dim paperFile As String
    paperFile = CurrentProject.path & "\paper.xlsx"
    
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "RawPaper", paperFile, True, "2019!"

End Sub

Public Sub FillWeightV3()

    Dim db     As DAO.Database
    Set db = CurrentDb
  
    db.Execute "CreateWeight", dbFailOnError
    Debug.Print "CreateWeight", db.RecordsAffected
    

    Dim rsPaper As Recordset
    Set rsPaper = db.OpenRecordset("Paper", dbOpenTable)

    Dim qd As DAO.QueryDef
    Set qd = db.QueryDefs("InsertWeight")

    Do While Not rsPaper.EOF
        Dim authors() As String
        authors = ExtractAUSAuthors(rsPaper!Address)
        If (Not Not authors) <> 0 Then
            For Each an In authors
                qd.Parameters("PaperID").Value = rsPaper!Id
                qd.Parameters("PaperTitle").Value = rsPaper!Title
                qd.Parameters("AuthorName").Value = an
                qd.Execute dbFailOnError
            Next an
        End If

        rsPaper.MoveNext
    Loop

End Sub

Public Function Foo()

    Beep
End Function

Public Function InitDataV3()
   
    'Test
    ImportSheetV3

    OpenFormV3
    Beep
End Function

Sub OpenTablesV3()
    DoCmd.OpenTable "College"
    DoCmd.OpenTable "Position"
End Sub

Sub ImportSheetV3()
    Set fso = New Scripting.FileSystemObject
    Dim imported As String
    imported = CurrentProject.path & "\imported.txt"
    
    If Not fso.FileExists(imported) Then
        CloseAllTablesV3
        DeleteAllRelationsV3
        DeleteAllTablesV3
        
        ImportJournalV3
        ImportBookV3
        
        ImportPaperV3
        
        ExportUnknownJournalV3
        ExportUnknownBookV3
        
        ImportAuthorV3

        DeleteMFileV3
        
        OpenTablesV3
    Else
        Debug.Print "No Import"
    End If
    
End Sub

Sub ExportUnknownJournalV3()
    Dim filepath As String
    filepath = CurrentProject.path & ".\UnknownJournal.xlsx"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownJournal", filepath, True, "NotFoundInPaper"
End Sub

Sub ExportUnknownBookV3()
    Dim filepath As String
    filepath = CurrentProject.path & ".\UnknownBook.xlsx"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownBook", filepath, True, "NotFoundInPaper"

End Sub

Sub ExportUnknownAuthorV3()
    Dim filepath As String
    filepath = CurrentProject.path & ".\UnknownAuthor.xlsx"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownAuthor", filepath, True, "NotFoundInPaper"

End Sub

Sub OpenFormV3()
    DoCmd.OpenForm "DepForm"

End Sub

Sub ImportJournalV3()
    Dim filepath As String
    filepath = CurrentProject.path & ".\Journal.xlsx"
    
    
    'Table names without spaces must NOT use single quotes
    ' Take care of the trailing "!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawSSCI", filepath, True, "SSCI!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawSCIE", filepath, True, "SCIE!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawESCI", filepath, True, "ESCI!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawAHCI", filepath, True, "AHCI!"

End Sub

Sub ImportBookV3()
    Dim filepath As String
    filepath = CurrentProject.path & ".\Book.xlsx"
    

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawBKCI-S", filepath, True, "BKCI-S!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawBKCI-SSH", filepath, True, "BKCI-SSH!"

End Sub

Sub ImportAuthorV3()

    FillWeightV3
    Dim filepath As String
    filepath = CurrentProject.path & ".\Author.xlsx"
    
    
    'Table names without spaces must NOT use single quotes
    ' Take care of the trailing "!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawAuthor", filepath, True, "Author!"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawPosition", filepath, True, "Position!"
    
    Dim db     As DAO.Database
    Set db = CurrentDb
    Dim tables() As String
    tables = Split("College,Department,Position,Author", ",")
    For Each t In tables
        db.Execute "Create" & t, dbFailOnError
        Debug.Print "Create" & t, db.RecordsAffected
        db.Execute "Insert" & t, dbFailOnError
        Debug.Print "Insert" & t, db.RecordsAffected
        
    Next t
    
    
    
    Dim td     As DAO.TableDef
    Dim f      As DAO.Field2
    Set td = db("Author")
    Set f = td.CreateField("FullName")
    'f.Expression = "[FirstName] +        ', ' + [LastName]"
    f.Expression = "[LastName] +        ', ' + [FirstName]"
    td.Fields.Append f
    
    ExportUnknownAuthorV3

    tables = Split("College,Department,Position,Author", ",")
    For Each t In tables

        db.Execute "InsertUnknown" & t, dbFailOnError
        Debug.Print "InsertUnknown" & t, db.RecordsAffected
    Next t
End Sub

Sub ImportPaperV3()
    Dim filepath As String
    filepath = CurrentProject.path & "\Paper.xlsx"
    'Table names without spaces must NOT use single quotes
    ' Take care of the trailing "!"
    
    ' Error The search key was not found in any record. (Error 3709)
    ' DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawPaper", CurrentProject.Path & "\Page error.xlxs", True, "Paper!"

    ' DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "RawPaper", filepath, True, "Paper!"
    
    Dim db     As DAO.Database
    Set db = CurrentDb
  
    db.Execute "CreatePaper", dbFailOnError
    Debug.Print "CreatePaper", db.RecordsAffected
    
    
    Dim qd As DAO.QueryDef
    Set qd = db.QueryDefs("InsertPaper")

    For y = 2018 To 2019
        DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "RawPaper", filepath, True, y & "!"
    
    
        qd.Parameters("Year").Value = y

        qd.Execute dbFailOnError
        Debug.Print "InsertPaper", db.RecordsAffected
        
        DoCmd.DeleteObject acTable, "RawPaper"
        Debug.Print "Delete RawPaper", db.RecordsAffected
    Next y
        
End Sub

Sub CloseAllTablesV3()

    For Each obj In CurrentData.AllTables
        If Left(obj.name, 4) <> "MSys" Then
            Debug.Print "Closing " & obj.name
            DoCmd.Close acTable, obj.name, acSaveNo
        End If
    Next
End Sub

Sub DeleteAllTablesV3()
    For Each obj In CurrentData.AllTables
        If Left(obj.name, 4) <> "MSys" Then
            Debug.Print "Deleting " & obj.name
            DoCmd.DeleteObject acTable, obj.name
        End If
    Next
End Sub

Sub DeleteAllRelationsV3()
    Dim obj    As Relation
    For Each obj In CurrentDb.Relations
        If Left(obj.name, 4) <> "MSys" Then
            Debug.Print "Deleting " & obj.name
            CurrentDb.Relations.Delete obj.name
        End If
    Next
End Sub

