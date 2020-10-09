# UniWoS 1.0 - 核心源码

## VBA 代码

###App API 库
```vb
'Attribute VB_Name = "App"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-1-9"

Option Compare Database
Option Explicit

Sub CloseTables()
    Dim obj As AccessObject
    For Each obj In CurrentData.AllTables
        If Left(obj.Name, 4) <> "MSys" Then
            'Debug.Print "Closing " & obj.Name
            DoCmd.Close acTable, obj.Name, acSaveNo
        End If
    Next
End Sub

Sub CloseQueries()
    Dim obj As AccessObject
    For Each obj In CurrentData.AllQueries
        If Left(obj.Name, 4) <> "MSys" Then
            'Debug.Print "Closing " & obj.Name
            DoCmd.Close acQuery, obj.Name, acSaveNo
        End If
    Next
End Sub

Sub DeleteTables()
    Dim obj As AccessObject
    For Each obj In CurrentData.AllTables
        If Left(obj.Name, 4) <> "MSys" Then
            Debug.Print "Deleting " & obj.Name
            DoCmd.DeleteObject acTable, obj.Name
        End If
    Next
End Sub

Sub DeleteRelations()
    Dim obj    As Relation
    For Each obj In CurrentDb.Relations
        If Left(obj.Name, 4) <> "MSys" Then
            Debug.Print "Deleting " & obj.Name
            CurrentDb.Relations.Delete obj.Name
        End If
    Next
End Sub

Sub ClearTables()
    CloseTables
    CloseQueries
    DeleteRelations
    DeleteTables
End Sub

Public Function CheckTable(ByVal tblName As String) As Boolean
    Dim iCount As Integer
    iCount = DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "' And Type In (1,4,6)")
    'Debug.Print tblName, iCount
    If iCount = 1 Then
        CheckTable = True
    End If
End Function

Public Function Execute(ByVal Query As String, ParamArray Params() As Variant) As Integer
    Dim db As Database
    ' Must! CurrentDb will new object
    Set db = CurrentDb
    If UBound(Params) >= 0 Then
        Dim qd As QueryDef
        Set qd = db.QueryDefs(Query)
        Dim i As Integer
        For i = 0 To UBound(Params) Step 2
            qd.Parameters(Params(i)).Value = Params(i + 1)
        Next
        qd.Execute dbFailOnError
        Log.i "Execute", Query, "RecordsAffected", qd.RecordsAffected
        Execute = qd.RecordsAffected
    Else
        db.Execute Query, dbFailOnError
        Log.i "Execute", Query, "RecordsAffected", db.RecordsAffected
        Execute = db.RecordsAffected
    End If
End Function


Public Sub DeleteTable(Table As String)
    If CheckTable(Table) Then
        DoCmd.Close acTable, Table
        DoCmd.DeleteObject acTable, Table
        Log.i "DeleteTable", Table
    End If
End Sub

Public Function CheckFields(ByVal Table, ByVal Fields As String) As Boolean
    Dim db As Database
    Set db = CurrentDb
    Dim rs As Recordset
    'dbOpenTable only applies on editable table
    'Set rs = db.OpenRecordset("LinkAuthor", dbOpenTable)
    Set rs = db.OpenRecordset(Table, dbOpenSnapshot)
    Dim sField As String
    Dim aField As Variant
    Dim dicField As New Scripting.Dictionary
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        dicField.Add Trim(rs.Fields(i).Name), rs.Fields(i).Value
    Next i
    aField = Split(Fields, ";")
    For i = 0 To UBound(aField)
        sField = Trim(aField(i))
        If Not dicField.Exists(sField) Then
            'Log.W "CheckFields", sField & " field not found", "Table", Table, "Fields", Fields
            CheckFields = False
            Exit Function
        End If
    Next i
    CheckFields = True
End Function


Public Sub OpenFile(File As String)
    Shell "explorer.exe " & File
End Sub
```

### 主程序

```vb
Attribute VB_Name = "Main"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-1-10"

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
            App.Execute sQuery, "PaperID", rsPaper!ID, "AuthorName", ""
        Else
            Dim sName As String
            Dim vAuthors As Variant
            vAuthors = Split(rsPaper!AuthorNames, ";")
            
            Dim i As Integer
            For i = 0 To UBound(vAuthors)
                sName = Paper.FixName(vAuthors(i))
                App.Execute sQuery, "PaperID", rsPaper!ID, "AuthorName", sName
            Next i
        End If
        ' Addresses
            Dim sName As String
            Dim vAuthors As Variant

            vAuthors = Paper.ExtractAuthors
            
            Dim i As Integer
            For i = 0 To UBound(vAuthors)
                sName = Paper.FixName(vAuthors(i))
                App.Execute sQuery, "PaperID", rsPaper!ID, "AuthorName", sName
            Next i

        rsPaper.MoveNext
    Loop
    
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
            qd.Parameters("PaperID").Value = rsPaper!ID
            qd.Parameters("AuthorName").Value = ""
            qd.Execute dbFailOnError
        Else
            Authors = Split(rsPaper!AuthorNames, ";")
            Dim ii As Integer
            For ii = 0 To UBound(Authors)
                an = Authors(ii)
                qd.Parameters("PaperID").Value = rsPaper!ID
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
```

### 作者操作库

```vb
'Attribute VB_Name = "Author"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-2-9"

Option Compare Database
Option Explicit

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

Public Function GetAuthorName(ByVal Name As String) As String
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
```

### 论文操作库

```vb
'Attribute VB_Name = "Paper"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-3-19"

Option Compare Database
Option Explicit

Public Function ExtractAuthorsFromAddrs(Addrs As String) As String()
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    Dim sAuthors As String
  Dim sUniversity As String
  sUniversity = ""
  
  iEndPos = InStr(Addrs, "] " & sUniversity)
    If iEndPos = 0 Then
        'Debug.Print "[Error]ExtractAuthorsFromAddrs No authors;Addrs:" & Addrs
        ExtractAuthorsFromAddrs = aAuthor
        
        Exit Function
    End If
    
    iStartPos = InStrRev(Addrs, "[", iEndPos)
    sAuthors = Mid(Addrs, iStartPos + 1, iEndPos - iStartPos - 1)
    aAuthor = Split(sAuthors, "; ") 
    ExtractAuthorsFromAddrs = aAuthor
End Function

Public Function ExtractAuthorsFromIDs(IDs As String) As String()
    If IDs = "" Then
        Debug.Print "[Error]ExtractAuthorsFromIDs empty; IDs:" & IDs
        Exit Function
    End If
    
    Dim aAuthor() As String
    aAuthor = Split(IDs, ";")
    Dim i As Integer

    Dim a, Name As String
    For i = 0 To UBound(aAuthor)
        a = aAuthor(i)
        If a <> "" Then
            Name = Split(a, "/")(0)
            aAuthor(i) = Trim(Name)
        End If
    Next i
    
    ExtractAuthorsFromIDs = aAuthor
End Function

Public Function ExtractAuthorsText(Addrs As String) As String
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    Dim sAuthors As String
    Dim sUniversity As String
    sUniversity = ""
    
    iEndPos = InStr(Addrs, "] " & sUniversity)
    If iEndPos = 0 Then
        'Debug.Print "[Error]ExtractAuthorsText No authors;Addrs:" & Addrs
        ExtractAuthorsText = ""
        
        Exit Function
    End If
    
    iStartPos = InStrRev(Addrs, "[", iEndPos)
    sAuthors = Mid(Addrs, iStartPos + 1, iEndPos - iStartPos - 1) 
    ExtractAuthorsText = sAuthors
End Function

Public Function ExtractAuthors(Addrs As String) As String()
    Dim sAuthors, aAuthor() As String
    sAuthors = ExtractAuthorsText(Addrs)
    aAuthor = Split(sAuthors, "; ")
    ExtractAuthors = aAuthor
End Function

Public Function CountAuthors(Addrs As String) As Integer

    Dim aAuthor() As String
    aAuthor = ExtractAuthors(Addrs)
    If (Not Not aAuthor) = 0 Then
        Debug.Print "[Error]CountAuthors Addrs:" & Addrs
        CountAuthors = 0
        Exit Function
    End If
    CountAuthors = UBound(aAuthor) + 1
    'Debug.Print CountAuthors
End Function

Public Function SerializeAuthorNames(Addrs As String, ResearcherIDs As String, ORCIDs As String) As String

    Dim aAuthor() As String
    aAuthor = ExtractAuthors(Addrs)
    If UBound(aAuthor) = -1 Then
        'Debug.Print "[Error]SerializeAuthorNames No authors;Addrs:" & Addrs
        Exit Function
    End If

    Dim names As String
    Dim i As Integer
    Dim sFixedName As String
    names = FixName(aAuthor(0))
    For i = 1 To UBound(aAuthor)
        sFixedName = FixName(aAuthor(i))
        sFixedName = FixNameWithIDs(sFixedName, ResearcherIDs)
        sFixedName = FixNameWithIDs(sFixedName, ORCIDs)
        names = names & ";" & sFixedName
    Next i

    SerializeAuthorNames = names
End Function

Public Function FixName(ByVal FullName As String) As String
    If FullName = "" Then
        'Debug.Print "[Error]FixName No Full name"
        FixName = FullName
        Exit Function
    End If

    Dim aFullName As Variant
    
    aFullName = Split(Trim(FullName), ",")

    If UBound(aFullName) = 0 Then
        'Debug.Print "FixName warning, " & FullName
        FixName = FullName
        'Debug.Print FixName
        Exit Function
    End If
    ' WoS naming style: Last Name, First Name
    Dim sFirstName, sLastName As String
    sFirstName = Split(Trim(aFullName(1)), " ")(0)
    
    sLastName = Trim(aFullName(0))
    If sLastName = "" Then
        'Debug.Print "[Error]FixName No Last name; FullName:" & FullName
        FixName = FullName
        Exit Function
    End If
    
    sLastName = Split(sLastName, " ")(0)

    FixName = sFirstName + " " + sLastName
    'Debug.Print FixName
End Function

Public Function GetAbbrName(AuthorName As Variant) As String

    If IsNull(AuthorName) Then
        GetAbbrName = ""
        Exit Function
    End If
    Dim sFirstName, sLastName As String
    sFirstName = Left(AuthorName, 1) + "."
    sLastName = Split(AuthorName, " ")(1)

    GetAbbrName = sFirstName & " " & sLastName
End Function

Public Function FixNameWithIDs(Abbr As String, IDs As String) As String
    '    If Abbr = "W. Abuzaid" Then
    '        Debug.Print "[Debug]FixNameWithIDs Abbr:" & Abbr & ", IDs:" & IDs
    '    End If
    
    If Mid(Abbr, 2, 1) <> "." Then
        FixNameWithIDs = Abbr
        Exit Function
    End If

    If IDs = "" Then
        FixNameWithIDs = Abbr
        'Debug.Print "[Warn]FixNameWithIDs empty; Abbr:" & Abbr
        Exit Function
    End If
    
    Dim aAuthor() As String
    aAuthor = ExtractAuthorsFromIDs(IDs)
    Dim a As String
    Dim i As Integer
    For i = 0 To UBound(aAuthor)
        a = aAuthor(i)
        If a <> "" Then
            a = FixName(a)
            If (Mid(a, 2, 1) <> ".") And (Left(a, 1) = Left(Abbr, 1)) And (Left(a, 1) <> ",") Then
                Dim lastName As String
                lastName = Split(Abbr, " ")(1)
                If InStr(a, lastName) <> 0 Then
                    'Debug.Print "[Trace]FixNameWithIDs fixed; Abbr:" & Abbr & ", IDs:" & IDs
                    FixNameWithIDs = a
                    Exit Function
                End If
            End If
        End If
    Next i
    
    'Debug.Print "[Warn]FixNameWithIDs unfixed; Abbr:" & Abbr & ", IDs:" & IDs
    FixNameWithIDs = Abbr
End Function

Public Function SelectAuthor(Addrs As String, Order As Integer, ResearcherIDs As String, ORCIDs As String) As Variant
    If Order > 9 Or Order < 1 Then
        'Debug.Print "SelectAuthor error, Order: " & CStr(Order)
        SelectAuthor = Null
        Exit Function
    End If

    Dim aAuthor() As String

    aAuthor = ExtractAuthorsFromAddrs(Addrs)
    If (Not Not aAuthor) = 0 Then
        'Debug.Print "[Error]SelectAuthor No Authors; Order: " & CStr(Order) & ", Addrs:" & Addrs
        SelectAuthor = Null
        Exit Function
    End If
    Dim iIndex As Integer
    iIndex = Order - 1

    If UBound(aAuthor) >= 9 Then
        Debug.Print "SelectAuthor warning, UBound >= " & CStr(UBound(aAuthor))
    End If

    If iIndex > UBound(aAuthor) Then
        'Debug.Print "SelectAuthor error, iIndex > " & CStr(UBound(aAuthor))
        SelectAuthor = Null
        Exit Function
    End If

    Dim fixedName As String
    fixedName = FixName(aAuthor(iIndex))
    fixedName = FixNameWithIDs(fixedName, ResearcherIDs)
    fixedName = FixNameWithIDs(fixedName, ORCIDs)

    SelectAuthor = fixedName
    'Debug.Print "[Trace]SelectAuthor fixedName:" & fixedName
End Function

Public Function GetFirstName(sFullName) As String
    Dim aFullName() As String

    aFullName = Split(sFullName, " ")
    If UBound(aFullName) = 0 Then
        Debug.Print "GetFirstName failed, " & sFullName
        GetFirstName = ""
        Exit Function
    End If
    GetFirstName = Trim(aFullName(0))
End Function

Public Function GetLastName(sFullName) As String
    Dim aFullName() As String
    aFullName = Split(sFullName, " ")
    If UBound(aFullName) > 0 Then
        GetLastName = Trim(aFullName(1))
    End If
End Function

Public Function GetWoSAuthorName(ByVal FullName As String) As String
    FullName = Trim(FullName)
    Dim sFirstName, sLastName As String

    Dim iPos As String
    iPos = InStrRev(FullName, " ")
    
    If iPos = 0 Then
        Log.W "GetWoSAuthorName", "No Space", "FullName", FullName
        GetWoSAuthorName = FullName
        Exit Function
    End If
    
    sFirstName = Mid(FullName, 1, iPos - 1)
    sLastName = Mid(FullName, iPos + 1, Len(FullName) - iPos)
    
    GetWoSAuthorName = Trim(sLastName) & ", " & Trim(sFirstName)
End Function
```

### 评分操作库

```vb
'Attribute VB_Name = "Weight"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-3-29"

Option Compare Database
Option Explicit

'iFalcC:  faculty Count
'iAuthC: author Count
'iPapInd: paper index
'iCurrInd: current index

Public Function CalcScore(iID As Variant, iPapInd As Integer, iCurrInd As Integer, iFacC As Integer, iAuthC As Integer) As Double
    If iAuthC = 0 Then
        'Debug.Print "[Error]CalcScore zero"
        Exit Function
    End If

    Dim bIsFac As Byte
    If IsNull(iID) Then
        bIsFac = False
    Else
        bIsFac = True
    End If

    Dim dScore As Double
    dScore = 0#

    If iFacC = 0 Then                            ' without falcuty
        If Not bIsFac Then
            dScore = 1 / iAuthC
        End If
    Else                                         ' with faculty
        If bIsFac Then
            dScore = 1 / iFacC
        End If
    End If
    
    If iCurrInd = 0 Or iPapInd = iCurrInd Then
        CalcScore = FormatNumber(dScore, 2)
    Else
        CalcScore = 0#
    End If
End Function


```

### 日志操作库

```vb
'Attribute VB_Name = "Log"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-4-9"

Option Compare Database
Option Explicit

Private Function Out(Tag, Func, Reason As String, ParamArray Vars() As Variant) As String
    Dim vVars() As Variant
    vVars = Vars(0)
    Dim sMes As String
 
    Dim sKey, sVal, sVars As String
    sVars = ";"
    Dim i As Integer
    For i = 0 To UBound(vVars) Step 2
        sKey = CStr(vVars(i))
        sVal = CStr(vVars(i + 1))
        sVars = sVars & " " & sKey & ":" & sVal
    Next i

    sMes = "[" & Tag & "]" & Func & " " & Reason & sVars
    Debug.Print sMes
    Out = sMes
End Function

Public Sub D(Func, Reason As String, ParamArray Vars() As Variant)
Dim sLog As String
  sLog = Out("D", Func, Reason, Vars)
End Sub

Public Sub i(Func, Reason As String, ParamArray Vars() As Variant)
  Dim sLog As String
  sLog = Out("I", Func, Reason, Vars)
End Sub

Public Sub W(Func, Reason As String, ParamArray Vars() As Variant)
   Dim sLog As String
  sLog = Out("W", Func, Reason, Vars)
End Sub

Public Sub T(Func, Reason As String, ParamArray Vars() As Variant)
   Dim sLog As String
  sLog = Out("T", Func, Reason, Vars)
End Sub

Public Sub E(Func, Reason As String, ParamArray Vars() As Variant)
    Dim sLog As String
    sLog = Out("E", Func, Reason, Vars)
    
    Dim rResult As VbMsgBoxResult
    rResult = MsgBox(Reason, vbYes, Func)
End Sub
```

### 配置操作库

```vb
'Attribute VB_Name = "Config"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-1-9"

Option Compare Database
Option Explicit

Public Property Get SettingPath() As String
    SettingPath = CurrentProject.Path + Consts.SETTINGS_FILE
End Property

Public Property Get IndexKey(ByVal Index As String, Year As Integer) As String
    IndexKey = Index + "-" + CStr(Year)
End Property

Public Property Get SheetPath(Section As String, Key As String) As String
    SheetPath = CurrentProject.Path + Consts.SHEETS_DIR + Val(Section, Key) 
End Property

Public Property Get Val(Section As String, Key As String) As String
    Val = Word.System.PrivateProfileString(SettingPath, Section, Key)
End Property

Public Property Let Val(Section As String, Key As String, Value As String)
    Word.System.PrivateProfileString(SettingPath, Section, Key) = Value
End Property

Public Property Get ExportFile() As String
    Dim sTime As String
    sTime = CStr(Now)
    sTime = Replace(Now, "/", "-")
    sTime = Replace(sTime, ":", ".")
    ExportFile = CurrentProject.Path & Consts.EXPORT_DIR & sTime & " - " & Consts.EXPORT_FILE
End Property
```

#### 工具库

```vb
'Attribute VB_Name = "Util"
'Attribute File_Author = "Huangjin Zhou"
'Attribute File_Email = "zhouhuangjing@hotmail.com"
'Attribute FIle_Created = "2020-1-9"

Option Compare Database
Option Explicit
 
Private lngRowNumber As Long
Private colPrimaryKeys As VBA.Collection

Private dblRunningSum As Double
Private colKeyRS As VBA.Collection

Public Function ResetRowNumber() As Boolean
    Set colPrimaryKeys = New VBA.Collection
    lngRowNumber = 0

    Set colKeyRS = New VBA.Collection
    dblRunningSum = 0
    ResetRowNumber = True
End Function

Public Function RowNumber(UniqueKeyVariant As Variant) As Long
    Dim lngTemp As Long

    On Error Resume Next
    lngTemp = colPrimaryKeys(CStr(UniqueKeyVariant))
    If Err.Number Then
        lngRowNumber = lngRowNumber + 1
        colPrimaryKeys.Add lngRowNumber, CStr(UniqueKeyVariant)
        lngTemp = lngRowNumber
    End If

    RowNumber = lngTemp
End Function

Public Function RunningSum(UniqueKeyVariant As Variant, varNum As Variant) As Double
    Dim dblTemp As Double

    On Error Resume Next
    dblTemp = colKeyRS(CStr(UniqueKeyVariant))
    If Err.Number Then
        dblRunningSum = dblRunningSum + CDbl(varNum)
        colPrimaryKeys.Add dblRunningSum, CStr(UniqueKeyVariant)
        dblTemp = dblRunningSum
    End If

    RunningSum = dblTemp
End Function

Public Function TokenNum(strValues As String, strDelim As String) As Double

    Dim arSplit As Variant
    arSplit = Split(strValues, strDelim)
    TokenNum = UBound(arSplit) + 1
End Function
```



## SQL

#### NewAuthor

```sql
INSERT INTO Author ( Code, FullName, AuthorName, AbbrName, JobID, DepartmentID ) Values ( Code, FullName, AuthorName, AbbrName, JobID, DepartmentID )
```

#### UpdateWeightByPaper

```sql
Update Weight
SET AuthorName = NewAuthor
WHERE PaperID = PID 
AND AuthorName = OldAuthor 
```

#### DeleteWeightByPaper

```sql
 Delete
FROM Weight
WHERE PaperID = PaperID 
```

#### InsertWeight

```sql
INSERT INTO [Weight] (PaperID, AuthorName, FullName, LastName, FirstName, MiddleName, FirstInitial, MiddleInitial)
Values (PaperID, AuthorName, FullName, LastName, FirstName, MiddleName, FirstInitial, MiddleInitial);
```

#### CreateAuthor

```sql
CREATE TABLE Author (
    ID COUNTER PRIMARY KEY,
    Code VARCHAR,
    FullName VARCHAR,
    AuthorName VARCHAR,
    AbbrName VARCHAR,
    JobID int,
    DepartmentID int,

    LastName VARCHAR,
    FirstName VARCHAR,
    FirstInitial VARCHAR,
    MiddleName VARCHAR,
    MiddleInitial VARCHAR
);
```

#### CreateCollege

```sql
CREATE TABLE College (ID COUNTER PRIMARY KEY, NAME VARCHAR, ORDER BYTE);
```

#### CreateDepartment

```sql
CREATE TABLE Department (
    [ID] INT,
    [NAME] VARCHAR,
    CollegeID INT,
    FOREIGN KEY (CollegeID) REFERENCES College(ID)
);
```

#### CreateJob

```sql
CREATE TABLE Job (
    ID COUNTER PRIMARY KEY,
    Title VARCHAR,
    Display VARCHAR,
    [Order] BYTE
);
```

#### CreatePaper

```sql
CREATE TABLE Paper (
    ID COUNTER PRIMARY KEY,
    WoSID VARCHAR,
    DOI VARCHAR,
    Title Memo,
    [Year] int,
    [Index] byte,
    Addresses Memo,
    AuthorNames Memo,
    AuthorCount int,
    FullNames VARCHAR
);
```

#### CreateWeight

```sql
CREATE TABLE [Weight] (
    ID COUNTER PRIMARY KEY,
    PaperID int,
    AuthorName VARCHAR
    FullName VARCHAR, 
    LastName VARCHAR, 
    FirstName VARCHAR, 
    MiddleName VARCHAR, 
    FirstInitial VARCHAR, 
    MiddleInitial VARCHAR
);
```

#### InsertUnknownJob

```sql
INSERT INTO
    Job (ID, Title, Display, [Order])
VALUES
    (0, 'Unknown', 'Unknown', 0)
```

#### InsertUnknownAuthor

```sql
INSERT INTO Author ( FirstName, LastName, PositionID, DepartmentID )
SELECT  GetFirstName(AuthorName) 
       ,GetLastName(AuthorName) 
       ,0 
       ,0
FROM 
(
	SELECT  distinct AuthorName
	FROM SelectUnknownAuthor
)
```

#### InsertUnknownCollege

```sql
INSERT INTO College (ID, [Name]) VALUES (0, 'Unknown')
```

#### InsertUnknownDepartment

```sql
INSERT INTO Department (ID, [Name], CollegeID) VALUES (0, 'Unknown', 0)
```

#### InsertOtherCollege

```sql
INSERT INTO
    College ([Name])
SELECT  top 1 "Other"
FROM College
WHERE not exists ( 
SELECT  *
FROM College
WHERE [name]='Others')  
```

#### InsertOtherDepartment

```sql
INSERT INTO Department ([ID], [Name], CollegeID)
SELECT  DISTINCT DepartmentID 
       ,DepartmentName 
       ,1
FROM ImportAuthor
WHERE DepartmentName not IN ( SELECT DISTINCT [Name] FROM Department )  
```
#### MakeImportAuthorIn

```sql
SELECT  DISTINCT GetAuthorName([Full Name])         AS AuthorName 
       ,ID                                          AS Code 
       ,[Full Name]                                 AS FullName 
       ,GetAbbrName(AuthorName)                     AS AbbrName 
       ,FixTitle([Job Title])                       AS JobTitle 
       ,Department                                  AS DepartmentID 
       ,ExtractInDepName([Department Description])  AS DepartmentName 
       ,ExtractInCollName([Department Description]) AS CollegeName into ImportAuthor
FROM LinkAuthor
WHERE ID Not IN ( SELECT distinct Code FROM Author)  
```

#### MakeImportAuthorOut

```sql
SELECT  DISTINCT GetAuthorName([Name]) AS AuthorName 
       ,ID                             AS Code 
       ,[Name]                         AS FullName 
       ,GetAbbrName(AuthorName)        AS AbbrName 
       ,FixTitle(Title)                AS JobTitle 
       ,ExtractOutDepID(Department)    AS DepartmentID 
       ,ExtractOutDepName(Department)  AS DepartmentName 
       ,ExtractOutCollName(Department) AS CollegeName 
       into ImportAuthor
FROM LinkAuthor
WHERE ID Not IN ( SELECT distinct Code FROM Author )  
```

#### InsertCollege

```sql
INSERT INTO
    College ([Name])
SELECT
    DISTINCT CollegeName
FROM
    ImportAuthor
WHERE
    CollegeName not IN (
        SELECT DISTINCT
            [Name]
        FROM
            College
    );
```

#### InsertDepartment

```sql
INSERT INTO Department ([ID], [Name], CollegeID)
SELECT  DISTINCT DepartmentID 
       ,DepartmentName 
       ,[College.ID]
FROM ImportAuthor
INNER JOIN College
ON ImportAuthor.CollegeName = College.[Name]
WHERE DepartmentName not IN ( SELECT DISTINCT [Name] FROM Department )  
```

#### InsertOtherDepartment

```sql
INSERT INTO Department ([ID], [Name], CollegeID)
SELECT  DISTINCT DepartmentID 
       ,DepartmentName 
       ,1
FROM ImportAuthor
WHERE DepartmentName not IN ( SELECT DISTINCT [Name] FROM Department )  
```

#### InsertJob

```sql
INSERT INTO
       Job (Title, Display, [Order])
SELECT
       DISTINCT JobTitle,
       JobTitle,
       1
FROM
      ImportAuthor 
WHERE JobTitle not IN ( SELECT DISTINCT Title FROM Job )  
```

#### InsertAuthor

```sql
INSERT INTO Author ( Code, FullName, AuthorName, AbbrName, JobID, DepartmentID, LastName, FirstName, FirstInitial, MiddleName, MiddleInitial )
SELECT  DISTINCT Code 
       ,FullName 
       ,AuthorName 
       ,AbbrName 
       ,Job.ID 
       ,DepartmentID 
       ,GetAuthorLastName(FullName)      AS LastName 
       ,GetAuthorFirstName(FullName)     AS FirstName 
       ,GetAuthorFirstInitial(FullName)  AS FirstInitial 
       ,GetAuthorMiddleName(FullName)    AS MiddleName 
       ,GetAuthorMiddleInitial(FullName) AS MiddleInitial
FROM 
(ImportAuthor
	INNER JOIN Job
	ON ImportAuthor.JobTitle = Job.Title 
)
```

#### MakeImportPaper

```sql
SELECT  [UT (Unique WOS ID)]                                    AS WosID
       ,DOI
       ,[Article Title]                                         AS Title
       ,GetYear([Publication Year],[Early Access Date])         AS [Year]
       ,CByte(LinkPaper.[Index])                                AS [Index]
       ,Addresses
       ,SerializeAuthorNames(Addresses,[Researcher Ids],ORCIDs) AS AuthorNames
       ,[Researcher Ids]                                        AS ResearcherIDs
       ,ORCIDs
       ,CountAuthors(Addresses)                                 AS AuthorCount
       ,ExtractAuthorsText(Addresses)                           AS FullNames into ImportPaper
FROM LinkPaper
WHERE [UT (Unique WOS ID)] not IN ( SELECT WosID FROM Paper ); 
```

#### InsertPaper

```sql
INSERT INTO Paper ( WoSID, DOI, Title, [Year], [Index], Addresses, AuthorNames, AuthorCount, FullNames )
SELECT  WoSID
       ,DOI
       ,Title
       ,[Year]
       ,[Index]
       ,Addresses
       ,AuthorNames
       ,AuthorCount
       ,FullNames
FROM ImportPaper
```

#### InsertWeight

```sql
INSERT INTO [Weight] (PaperID, AuthorName, FullName, LastName, FirstName, MiddleName, FirstInitial, MiddleInitial)
Values (PaperID, AuthorName, FullName, LastName, FirstName, MiddleName, FirstInitial, MiddleInitial);
```

#### MakeUnknownAuthor

```sql
SELECT  DISTINCT SelectWeight.PaperID INTO UnknownAuthor
FROM SelectWeight
WHERE (((IsNull([AuthorID]))<>False));  
```