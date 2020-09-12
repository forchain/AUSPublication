Attribute VB_Name = "App"
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

Public Function Execute(Query As String, ParamArray Params() As Variant) As Integer
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











