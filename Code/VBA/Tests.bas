Attribute VB_Name = "Tests"
Option Compare Database
Option Explicit

Public Function TestLoadConfig()

End Function

Public Sub TestCurrentProject()

    Debug.Print CurrentProject.path


End Sub

Public Sub TestWordApplication()
    Dim ws As Word.System
    Set ws = Word.System                         ' create the Word application object
    
    Dim path As String
    path = CurrentProject.path + "/settings.ini"
    
    
    Debug.Print path
    Dim Index As String
    
    Index = ws.PrivateProfileString(path, "Index", "AHCI-2018")
    Debug.Print Index
    'Application.CurrentProject(    CurrentProject.Path "settings.ini"
    
    Index = ws.PrivateProfileString(path, "Staff", "FacultyDeparting")
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

Public Sub TestViewPaper()

    App.ClearTables

    Paper.ViewPaper
    
End Sub


Public Sub TestFillWeight()

    App.ClearTables

    Paper.ImportPaper
    
    Author.ImportAuthor
    
    Weight.FillWeight
    
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
    

    Dim t As String
    For i = 1 To 9
        t = "UpdateKnownAuthor" & CStr(i)
        CurrentDb.Execute t, dbFailOnError
        Debug.Print t, CurrentDb.RecordsAffected
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

