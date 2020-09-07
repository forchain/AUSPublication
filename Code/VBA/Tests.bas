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
    Print
    Dim Index As String
    
    Index = ws.PrivateProfileString(path, "Index", "AHCI-2018")
    Debug.Print Indexp
    'Application.CurrentProject(    CurrentProject.Path "settings.ini"
    
    Index = ws.PrivateProfileString(path, "Staff", "FacultyDeparting")
    Debug.Print Index
    
    
End Sub

Public Sub TestImportPaper()

    App.ClearTables

    Dim currYear As Integer
    currYear = Year(Date)

    Dim y As Integer
    Dim i As Integer
    Dim key As String
    Dim path As String
    
    CurrentDb.Execute "CreatePaper", dbFailOnError
    Debug.Print "CreatePaper", CurrentDb.RecordsAffected

    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.QueryDefs("InsertPaper")
    
    For y = Consts.BEIGN_YEAR To currYear
        For i = 0 To UBound(Consts.INDICES) - 1
            
            key = Config.IndexKey(Consts.INDICES(i), y)

            path = Config.SheetPath(key)
            
            Debug.Print path

            DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, "RawPaper", path, True, Consts.SHEET_PAPER & "!"
            
            qd.Parameters("Year").Value = y
            qd.Parameters("Index").Value = i + 1

            qd.Execute dbFailOnError
            Debug.Print "InsertPaper", CurrentDb.RecordsAffected
        
            DoCmd.DeleteObject acTable, "RawPaper"
            Debug.Print "Delete RawPaper", CurrentDb.RecordsAffected
        Next i
    Next y

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



