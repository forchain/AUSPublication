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
        authors = Paper.ExtractAuthors(rsPaper!Addresses)
        If UBound(authors) >= 0 Then
            Dim iI As Integer
            For iI = 0 To UBound(authors)
                an = authors(iI)
                qd.Parameters("PaperID").Value = rsPaper!Id
                qd.Parameters("AuthorName").Value = Paper.FixName(an)
                qd.Execute dbFailOnError
            Next iI
        Else
            qd.Parameters("PaperID").Value = rsPaper!Id
            qd.Parameters("AuthorName").Value = ""
            qd.Execute dbFailOnError
        End If

        rsPaper.MoveNext
    Loop
    
End Sub


'iFalcC:  faculty Count
'iAuthC: author Count
'iPapInd: paper index
'iCurrInd: current index

Public Function CalcScore(iID As Variant, iPapInd As Integer, iCurrInd As Integer, iFacC As Integer, iAuthC As Integer) As Double

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
        CalcScore = dScore
    Else
        CalcScore = 0#
    End If

End Function









