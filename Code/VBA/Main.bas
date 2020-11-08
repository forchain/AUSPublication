Attribute VB_Name = "Main"

Option Compare Database
Option Explicit

Public Sub UpdateWeightByPaper(ByVal PaperID As Integer, ByVal OldNames As String, ByVal NewNames As String)

    Dim sAuthorName As Variant
   
    Dim aOldName, aNewName As Variant
    
    aOldName = Split(OldNames, ";")
    aNewName = Split(NewNames, ";")
    
    If UBound(aOldName) = UBound(aNewName) Then
        Dim i As Integer
        Dim sNewAuthor, sOldAuthor As String
        
        For i = 0 To UBound(aNewName)
            sNewAuthor = Trim(aNewName(i))
            sOldAuthor = Trim(aOldName(i))
            If sNewAuthor <> sOldAuthor Then
                App.Execute "UpdateWeightByPaper", "PID", PaperID, "OldAuthor", sOldAuthor, "NewAuthor", sNewAuthor

            End If
        Next

    Else
        App.Execute "DeleteWeightByPaper", "PID", PaperID

        For Each sAuthorName In aNewName
            App.Execute "InsertWeight", "PaperID", PaperID, "AuthorName", sAuthorName
        Next
    End If
   

End Sub

Public Sub CreateTables()
    Dim dicUnknown As New Scripting.Dictionary
    Dim dicOther As New Scripting.Dictionary
    Dim dicDefault As New Scripting.Dictionary
    
    dicUnknown.Add "Author", False
    dicUnknown.Add "College", True
    dicUnknown.Add "Department", True
    dicUnknown.Add "Job", True
    dicUnknown.Add "Paper", False

    dicUnknown.Add "ImportScore", False
    dicUnknown.Add "Score", False
    dicUnknown.Add "Match", False
    dicUnknown.Add "Setting", False
    
    dicOther.Add "College", True
    dicDefault.Add "Setting", True


    Dim bUnknown, bOthers As Boolean
    Dim sKey, sQuery As String

    For Each sKey In dicUnknown.Keys
        If Not CheckTable(sKey) Then
            sQuery = "Create" + sKey
            App.Execute sQuery
            bUnknown = dicUnknown.Item(sKey)
            If bUnknown Then
                sQuery = "InsertUnknown" + sKey
                App.Execute sQuery
            End If
            If dicOther.Exists(sKey) Then
                sQuery = "InsertOther" + sKey
                App.Execute sQuery
            End If
            If dicDefault.Exists(sKey) Then
                sQuery = "InsertDefault" + sKey
                App.Execute sQuery
            End If
        End If
    Next sKey

End Sub

Public Sub CreateFields()

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim field As DAO.Field2

    Set db = CurrentDb()
    Set td = db.TableDefs("Match")

    Set field = td.CreateField("MinimumMatched", dbBoolean)
    field.Expression = "IsMinimumMatched(AuthorID)"
    td.Fields.Append field

    Set field = td.CreateField("FirstNameMatched", dbBoolean)
    field.Expression = "IsFirstNameMatched(FirstNameCheck, PaperFirstName, AuthorFirstName )"
    td.Fields.Append field

    Set field = td.CreateField("MiddleNameMatched", dbBoolean)
    field.Expression = "IsMiddleNameMatched(MiddleNameCheck, PaperMiddleName, AuthorMiddleName )"
    td.Fields.Append field

    Set field = td.CreateField("MiddleInitialMatched", dbBoolean)
    field.Expression = "IsMiddleInitialMatched(MiddleInitialCheck, PaperMiddleInitial, AuthorMiddleInitial)"
    td.Fields.Append field
    
        Set field = td.CreateField("AllMatched", dbBoolean)
    field.Expression = "MinimumMatched and FirstNameMatched and MiddleNameMatched and MiddleInitialMatched"
    td.Fields.Append field

End Sub

Public Function ImportAuthor(EmplType As Byte, ByVal Path As String) As Integer
    Dim sFunc As String
    sFunc = "ImportAuthor"
    
    App.DeleteTable "LinkAuthor"
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "LinkAuthor", Path, True
        
    Dim i As Integer
    Dim sQuery As String
    
    App.DeleteTable "ImportAuthor"
    Dim sInFields, sOutFields, sStudentFields As String
    sInFields = "Type;Full Name;ID;Current Hire Date;Job Title;Department;Department Description"
    sOutFields = "Empl Type;Department;Name;ID;Current Hire Date;Termination Date;Title"
    sStudentFields = "Last Name;ID;First Name;College/ School;Department;Enrollment Year;Graduation Year;Position"
    
    If App.CheckFields("LinkAuthor", sInFields) Then
        sQuery = "MakeImportAuthorIn"
    ElseIf App.CheckFields("LinkAuthor", sOutFields) Then
        sQuery = "MakeImportAuthorOut"
    ElseIf App.CheckFields("LinkAuthor", sStudentFields) Then
        sQuery = "MakeImportAuthorStudent"
    Else
        Log.E sFunc, "Invalid fields", "Path", Path
        Exit Function
    End If
    
    Dim iRows As Integer
    iRows = App.Execute(sQuery)

    If iRows = 0 Then
        MsgBox "No new records imported", Title:="Import"
        ImportAuthor = iRows
        Exit Function
    End If
    
    '1 Faculty; 2 Staff
    If EmplType = 1 Then
        sQuery = "InsertCollege"
        App.Execute sQuery
        sQuery = "InsertDepartment"
        App.Execute sQuery
    Else
        sQuery = "InsertOtherDepartment"
        App.Execute sQuery
    End If
    
    sQuery = "InsertJob"
    App.Execute sQuery
    
    sQuery = "InsertAuthor"
    App.Execute sQuery

    
    App.DeleteTable "ImportMatch"
    sQuery = "MakeImportMatchByAuthor"
    App.Execute sQuery
    'replace empty records in match with those matched in ImportMatch
    sQuery = "DeleteUnknownMatch"
    App.Execute sQuery
    sQuery = "InsertMatch"

    App.Execute sQuery, _
                "FirstNameCheck", Config.Setting("FirstNameCheck"), _
                "MiddleNameCheck", Config.Setting("MiddleNameCheck"), _
                "MiddleInitialCheck", Config.Setting("MiddleInitialCheck")
    
    App.DeleteTable "ResolvedMatch"
    sQuery = "MakeResolvedMatch"
    App.Execute sQuery
    
    
    sQuery = "UpdateScore"
    App.Execute sQuery
    
    'Log.i sFunc, "Imported", "iRows", iRows
    MsgBox iRows & " records imported", Title:="Import"

    
    ImportAuthor = iRows
End Function

Public Function ImportPaper(ByVal Index As Integer, ByVal Path As String) As Integer
    Dim sFunc As String
    sFunc = "ImportAuthor"
    
    Dim sLinkTable As String
    sLinkTable = "LinkPaper"

    App.DeleteTable sLinkTable

    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, sLinkTable, Path, True
        
    Dim sFields, sOutFields As String
    sFields = "Publication Type;Authors;Book Authors;Book Editors;Book Group Authors;Author Full Names;Book Author Full Names;Group Authors;Article Title;Source Title;Book Series Title;Book Series Subtitle;Language;Document Type;Conference Title;Conference Date;Conference Location;Conference Sponsor;Conference Host;Author Keywords;Keywords Plus;Abstract;Addresses;Reprint Addresses;Email Addresses;Researcher Ids;ORCIDs;Funding Orgs;Funding Text;Cited References;Cited Reference Count;Times Cited, WoS Core;Times Cited, All Databases;180 Day Usage Count;Since 2013 Usage Count;Publisher;Publisher City;Publisher Address;ISSN;eISSN;ISBN;Journal Abbreviation;Journal ISO Abbreviation;Publication Date;Publication Year;Volume;Issue;Part Number;Supplement;Special Issue;Meeting Abstract;Start Page;End Page;Article Number;DOI;Book DOI;Early Access Date;Number of Pages;WoS Categories;Research Areas;IDS Number;UT (Unique WOS ID);Pubmed Id;Open Access Designations;Highly Cited Status;Hot Paper Status;Date of Export"
    
    If Not App.CheckFields(sLinkTable, sFields) Then
        Log.E "ImportPaper", "Invalid fields", "Index", Index, "Path", Path, "sFields", sFields
        Exit Function
    End If
    
    App.DeleteTable "ImportPaper"
    
    Dim sQuery As String
    Dim iRows As Integer
    
    sQuery = "MakeImportPaper"
    iRows = App.Execute(sQuery, sLinkTable & ".Index", Index)

    If iRows = 0 Then
        MsgBox "No new records imported", Title:="Import"
        ImportPaper = iRows
        Exit Function
    End If
    
    sQuery = "InsertPaper"
    App.Execute sQuery
    
    
    App.Execute "ClearImportScore"
    Dim rsPaper As Recordset
    Set rsPaper = CurrentDb.OpenRecordset("SelectImportPaper", dbOpenSnapshot)
    
    Dim i As Integer
    Dim sName As String
    sQuery = "InsertImportScore"
    Do While Not rsPaper.EOF
        Dim vAuthors As Variant
        vAuthors = Split(rsPaper!FullNames, ";")

        For i = 0 To UBound(vAuthors)
            sName = Trim(vAuthors(i))
            App.Execute sQuery, "PaperID", rsPaper!ID, "WoSID", rsPaper!WoSID, "Index", rsPaper![Index], "AuthorCount", UBound(vAuthors) + 1, "FullName", sName, "LastName", Paper.GetLastName(sName), "FirstName", Paper.GetFirstName(sName), "MiddleName", Paper.GetMiddleName(sName), "FirstInitial", Paper.GetFirstInitial(sName), "MiddleInitial", Paper.GetMiddleInitial(sName)
        Next i
        rsPaper.MoveNext
    Loop
    
    sQuery = "InsertScore"
    App.Execute sQuery

    App.DeleteTable "ImportMatch"
    sQuery = "MakeImportMatchByPaper"
    App.Execute sQuery
    sQuery = "InsertMatch"
    App.Execute sQuery, _
                "FirstNameCheck", Config.Setting("FirstNameCheck"), _
                "MiddleNameCheck", Config.Setting("MiddleNameCheck"), _
                "MiddleInitialCheck", Config.Setting("MiddleInitialCheck")
    
    App.DeleteTable "ResolvedMatch"
    sQuery = "MakeResolvedMatch"
    App.Execute sQuery
    
    sQuery = "UpdateScore"
    App.Execute sQuery
    
    'Log.i sFunc, "Imported", "iRows", iRows
    MsgBox iRows & " records imported", Title:="Import"
    ImportPaper = iRows

End Function

Public Sub FillWeight()
    
    Dim db     As DAO.Database
    Set db = CurrentDb
  
    db.Execute "CreateWeight", dbFailOnError
    Debug.Print "CreateWeight", db.RecordsAffected
    

    Dim rsPaper As Recordset
    Set rsPaper = db.OpenRecordset("Paper", dbOpenTable)
    '    Set rsPaper = db.OpenRecordset("ThanViewPaper", dbOpenDynaset)

    Dim qd As DAO.QueryDef
    Dim an As String
    Set qd = db.QueryDefs("InsertWeight")

    Do While Not rsPaper.EOF
        Dim Authors() As String
        '        If rsPaper!Id = 828 Then
        '            Debug.Print rsPaper!AuthorNames
        '        End Ifn
       
        If IsNull(rsPaper!FullNames) Then
            qd.Parameters("PaperID").Value = rsPaper!ID
            qd.Parameters("AuthorName").Value = ""
            
            qd.Parameters("FullName").Value = ""
            qd.Parameters("FirstName").Value = ""
            qd.Parameters("LastName").Value = ""
            qd.Parameters("MiddleName").Value = ""
                
            qd.Parameters("FristInitial").Value = ""
            qd.Parameters("MiddleInitial").Value = ""
            qd.Execute dbFailOnError
        Else
            Authors = Split(rsPaper!FullNames, ";")
            Dim ii As Integer
            For ii = 0 To UBound(Authors)
                an = Authors(ii)
                qd.Parameters("PaperID").Value = rsPaper!ID
                qd.Parameters("AuthorName").Value = Paper.FixName(an)
                qd.Parameters("FullName").Value = an
                qd.Parameters("FirstName").Value = Paper.GetFirstName(an)
                qd.Parameters("LastName").Value = Paper.GetLastName(an)
                qd.Parameters("MiddleName").Value = Paper.GetMiddleName(an)
                
                qd.Parameters("FristInitial").Value = Paper.GetFirstInitial(an)
                qd.Parameters("MiddleInitial").Value = Paper.GetMiddleInitial(an)
                
                qd.Execute dbFailOnError
            Next ii
        End If

        rsPaper.MoveNext
    Loop
    
    ' Unknown Author
    '    Dim sPath As String
    '    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_UNKNOWN_AUTHOR_FILE)
    '
    '    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownAuthor", sPath, True, Consts.SHEET_UNKNOWN_AUTHOR
End Sub

Public Sub MakeUnknownAuthor()
    App.DeleteTable "UnknownAuthor"
    
    App.Execute "MakeUnknownAuthor"
End Sub

Public Function ImportSheets()
    App.ClearTables

    ImportPaper
    
    ImportAuthor
    
    FillWeight
    
    MakeUnknownAuthor
    
    Dim rResult As VbMsgBoxResult
    rResult = MsgBox("Done", vbYes, "Import Sheets")
End Function

