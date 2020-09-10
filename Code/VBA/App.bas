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

Public Function CheckTable(tblName As String) As Boolean

    If DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "' And Type In (1,4,6)") = 1 Then

        CheckTable = True

    End If

End Function
