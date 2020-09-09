Attribute VB_Name = "Weight"
Option Compare Database
Option Explicit

Public Sub FillWeight()
    
    Dim db     As DAO.Database
    Set db = CurrentDb
  
    db.Execute "CreateWeight", dbFailOnError
    Debug.Print "CreateWeight", db.RecordsAffected
    

    Dim rsPaper As Recordset
    'Set rsPaper = db.OpenRecordset("Paper", dbOpenTable)
    Set rsPaper = db.OpenRecordset("ViewPaper", dbOpenDynaset)

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
    Dim sPath As String
    sPath = Config.SheetPath(Consts.SECTION_AUTHOR, Consts.KEY_UNKNOWN_AUTHOR_FILE)
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SelectUnknownAuthor", sPath, True, Consts.SHEET_UNKNOWN_AUTHOR
End Sub


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










