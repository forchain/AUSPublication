Attribute VB_Name = "Author"
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
    CurrentDb.Execute "CreateCollege", dbFailOnError
    Debug.Print "CreateCollege", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertUnknownCollege", dbFailOnError
    Debug.Print "InsertUnknownCollege", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertCollege", dbFailOnError
    Debug.Print "InsertCollege", CurrentDb.RecordsAffected
    
    ' Department
    CurrentDb.Execute "CreateDepartment", dbFailOnError
    Debug.Print "CreateDepartment", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertUnknownDepartment", dbFailOnError
    Debug.Print "InsertUnknownDepartment", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertDepartment", dbFailOnError
    Debug.Print "InsertDepartment", CurrentDb.RecordsAffected
    
    ' Job
    
    CurrentDb.Execute "CreateJob", dbFailOnError
    Debug.Print "CreateJob", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertUnknownJob", dbFailOnError
    Debug.Print "InsertUnknownJob", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertJob", dbFailOnError
    Debug.Print "InsertJob", CurrentDb.RecordsAffected
    
    Dim path As String
    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_JOB_FILE)
    sSheet = Consts.SHEET_JOB & "!"
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Job", sPath, True, Consts.SHEET_JOB
    
    DoCmd.DeleteObject acTable, "Job"
    Debug.Print "Delete Job", CurrentDb.RecordsAffected
    
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "Job", sPath, True, sSheet
    
    ' Author
    
    CurrentDb.Execute "CreateAuthor", dbFailOnError
    Debug.Print "CreateAuthor", CurrentDb.RecordsAffected
    
    
    CurrentDb.Execute "InsertAuthor", dbFailOnError
    Debug.Print "InsertAuthor", CurrentDb.RecordsAffected
   
    'DoCmd.OpenTable "Author"
    'Debug.Print "TestImportAuthor"
End Sub

Public Function ExtractOutDepID(Dep As String) As Integer
    Dim aToken
    aToken = Split(Dep, "-")
    ExtractOutDepID = CInt(Trim(aToken(0)))
End Function

Public Function ExtractOutCollName(Dep As String) As String
    'Debug.Print "[Debug]ExtractOutCollName Dep:" & Dep
    Dim aToken() As String
    aToken = Split(Dep, "-")
    
    ExtractOutCollName = ExtractInCollName(aToken(1))
End Function

Public Function ExtractOutDepName(Dep As String) As String

    'Debug.Print "[Debug]ExtractOutDepName Dep:" & Dep
    Dim aToken() As String
    aToken = Split(Dep, "-")
    
    ExtractOutDepName = ExtractInDepName(aToken(1))
End Function

Public Function ExtractInCollName(Dep As String) As String
    Dim aToken
    
    aToken = Split(Dep, "/")
    
    ExtractInCollName = Trim(aToken(0))
End Function

Public Function ExtractInDepName(Dep As String) As String
    Dim aToken
    
    aToken = Split(Dep, "/")
    
    ExtractInDepName = Trim(aToken(1))
End Function

Public Function GetAuthorName(Name As String) As String
    Dim aToken
    aToken = Split(Name, " ")
    
    Dim sFirstName As String
    sFirstName = Trim(aToken(0))
    
    Dim sLastName As String
    sLastName = Trim(aToken(UBound(aToken)))
    
    GetAuthorName = sFirstName + " " + sLastName
    
End Function

Public Function FixTitle(Title As String) As String
    Dim aToken() As String
    Dim sTitle As String
    
    sTitle = Replace(Title, "-Prof", " Prof")
    
    aToken = Split(sTitle, "-")
    
    sTitle = Trim(aToken(0))

    
    
    aToken = Split(sTitle, " of")
    
    sTitle = Trim(aToken(0))
    
    '    aToken = Split(sTitle, ".")
    '
    '    sTitle = Trim(aToken(0))
    
    aToken = Split(sTitle, "Prof")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Professor"
    End If
    
    aToken = Split(sTitle, "Inst")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Instructor"
    End If
    
    
    aToken = Split(sTitle, "Lecturer")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Lecturer"
    End If
    
    aToken = Split(sTitle, "Dean")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Dean"
    End If
    
    
    FixTitle = Trim(sTitle)

End Function

