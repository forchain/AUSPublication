Attribute VB_Name = "Main"
Option Compare Database
Option Explicit

Public Sub CreateTables()
    Dim dicUnknown As New Scripting.Dictionary
    Dim dicOther As New Scripting.Dictionary
    
    dicUnknown.Add "Author", False
    dicUnknown.Add "College", True
    dicUnknown.Add "Department", True
    dicUnknown.Add "Job", True
    dicUnknown.Add "Paper", False
    dicUnknown.Add "Weight", False
    
    dicOther.Add "College", True

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
        End If
    Next sKey

End Sub

Public Function ImportAuthor(EmplType As Byte, ByVal Path As String) As Integer
    Dim sFunc As String
    sFunc = "ImportAuthor"
    
    App.DeleteTable "LinkAuthor"
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "LinkAuthor", Path, True
        
    Dim i As Integer
    Dim sQuery As String
    
    Dim sInFields, sOutFields As String
    sInFields = "Type;Full Name;ID;Current Hire Date;Job Title;Department;Department Description"
    sOutFields = "Empl Type;Department;Name;ID;Current Hire Date;Termination Date;Title"
    
    If App.CheckFields("LinkAuthor", sInFields) Then
        sQuery = "MakeImportAuthorIn"
    ElseIf App.CheckFields("LinkAuthor", sOutFields) Then
        sQuery = "MakeImportAuthorOut"
    Else
        Log.E sFunc, "Invalid fields", "Path", Path
        Exit Function
    End If
    
    App.DeleteTable "ImportAuthor"
    
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
        App.Execute (sQuery)
        sQuery = "InsertDepartment"
        App.Execute (sQuery)
    Else
        sQuery = "InsertOtherDepartment"
        App.Execute (sQuery)
    End If
    
    sQuery = "InsertJob"
    App.Execute (sQuery)
    
    sQuery = "InsertAuthor"
    App.Execute (sQuery)
    
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
    
    Dim rsPaper As Recordset
    Set rsPaper = CurrentDb.OpenRecordset("SelectImportPaper", dbOpenSnapshot)

    sQuery = "InsertWeight"
    Do While Not rsPaper.EOF
        If IsNull(rsPaper!AuthorNames) Then
            App.Execute sQuery, "PaperID", rsPaper!Id, "AuthorName", ""
        Else
            Dim sName As String
            Dim vAuthors As Variant
            vAuthors = Split(rsPaper!AuthorNames, ";")
            
            Dim i As Integer
            For i = 0 To UBound(vAuthors)
                sName = Paper.FixName(vAuthors(i))
                App.Execute sQuery, "PaperID", rsPaper!Id, "AuthorName", sName
            Next i
        End If

        rsPaper.MoveNext
    Loop
    
    'Log.i sFunc, "Imported", "iRows", iRows
    
    ImportPaper = iRows

End Function

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
        Dim Authors() As String
        '        If rsPaper!Id = 828 Then
        '            Debug.Print rsPaper!AuthorNames
        '        End If
       
        If IsNull(rsPaper!AuthorNames) Then
            qd.Parameters("PaperID").Value = rsPaper!Id
            qd.Parameters("AuthorName").Value = ""
            qd.Execute dbFailOnError
        Else
            Authors = Split(rsPaper!AuthorNames, ";")
            Dim ii As Integer
            For ii = 0 To UBound(Authors)
                an = Authors(ii)
                qd.Parameters("PaperID").Value = rsPaper!Id
                qd.Parameters("AuthorName").Value = Paper.FixName(an)
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

    If App.CheckTable("UnknownAuthor") Then
        DoCmd.DeleteObject acTable, "UnknownAuthor"
        Debug.Print "Delete UnknownAuthor", CurrentDb.RecordsAffected
    End If
    CurrentDb.Execute "MakeUnknownAuthor", dbFailOnError
    Debug.Print "MakeUnknownAuthor", CurrentDb.RecordsAffected
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

