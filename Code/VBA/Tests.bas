Attribute VB_Name = "Tests"

Option Compare Database
Option Explicit

Public Function TestLoadConfig()

End Function

Public Sub TestCurrentProject()

    Debug.Print CurrentProject.Path


End Sub

Sub TestOpenFile()
    Dim Shex As Object
    Set Shex = CreateObject("Shell.Application")
    tgtfile = "C:\Nax\dud.txt"
    Shex.Open (tgtfile)
End Sub

Public Sub TestWordApplication()
    Dim ws As Word.System
    Set ws = Word.System                         ' create the Word application object
    
    Dim Path As String
    Path = CurrentProject.Path + "/settings.ini"
    
    
    Debug.Print Path
    Dim Index As String
    
    Index = ws.PrivateProfileString(Path, "Index", "AHCI-2018")
    Debug.Print Index
    'Application.CurrentProject(    CurrentProject.Path "settings.ini"
    
    Index = ws.PrivateProfileString(Path, "Staff", "FacultyDeparting")
    Debug.Print Index
    
    
End Sub

Public Sub TestImportWeight()

    App.ClearTables

    Author.ImportAuthor
    DoCmd.OpenTable "Author"

    Paper.ImportPaper
    DoCmd.OpenTable "Paper"

End Sub

Public Sub TestImportAuthor()

    App.ClearTables

    Author.ImportAuthor
    
    DoCmd.OpenTable "Author"

End Sub

Public Sub TestImportPaper()

    App.ClearTables

    Paper.ImportPaper
    DoCmd.OpenTable "Paper"

End Sub

Public Sub TestAppendPaper()

    App.CloseTables
    App.CloseQueries

    Paper.ImportPaper
    DoCmd.OpenTable "Paper"

End Sub

Public Sub TestViewAuthor()

    App.ClearTables

    Paper.ViewPaper
    
End Sub

Public Sub TestFixAbbr()
    CurrentDb.Execute "CreateAbbr", dbFailOnError
    Debug.Print "CreateAbbr", CurrentDb.RecordsAffected
    
    CurrentDb.Execute "InsertAbbr", dbFailOnError
    Debug.Print "InsertAbbr", CurrentDb.RecordsAffected
    

    Dim T As String
    For i = 1 To 9
        T = "UpdateKnownAuthor" & CStr(i)
        CurrentDb.Execute T, dbFailOnError
        Debug.Print T, CurrentDb.RecordsAffected
    Next i
End Sub

Public Sub TestYear()
    Dim myDate As Date

    myDate = Date
    Debug.Print myDate
    Debug.Print TypeName(myDate)
    Debug.Print Year(myDate)
    Debug.Print TypeName(Year(Date))
End Sub

Public Sub TestConfig()
    Debug.Print Config.SettingPath
 

    Config.Val("TestSection", "TestKey") = "TestValue"
 
    Debug.Print Config.Val("TestSection", "TestKey")
End Sub


