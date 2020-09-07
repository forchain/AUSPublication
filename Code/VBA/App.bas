Attribute VB_Name = "App"
Option Compare Database
Option Explicit

Sub CloseTables()
    Dim obj As AccessObject
    For Each obj In CurrentData.AllTables
        If Left(obj.name, 4) <> "MSys" Then
            Debug.Print "Closing " & obj.name
            DoCmd.Close acTable, obj.name, acSaveNo
        End If
    Next
End Sub

Sub DeleteTables()
    Dim obj As AccessObject
    For Each obj In CurrentData.AllTables
        If Left(obj.name, 4) <> "MSys" Then
            Debug.Print "Deleting " & obj.name
            DoCmd.DeleteObject acTable, obj.name
        End If
    Next
End Sub

Sub DeleteRelations()
    Dim obj    As Relation
    For Each obj In CurrentDb.Relations
        If Left(obj.name, 4) <> "MSys" Then
            Debug.Print "Deleting " & obj.name
            CurrentDb.Relations.Delete obj.name
        End If
    Next
End Sub

Sub ClearTables()
    CloseTables
    DeleteTables
    DeleteRelations
End Sub
    


