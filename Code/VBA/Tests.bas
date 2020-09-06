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
    Debug.Print Index
    'Application.CurrentProject(    CurrentProject.Path "settings.ini"
    
    Index = ws.PrivateProfileString(path, "Staff", "FacultyDeparting")
    Debug.Print Index
    
    
End Sub

Public Sub TestImportPapers()

    Dim currYear As Integer
    currYear = Year(Date)

    Dim y As Integer
    Dim i As Integer
    Dim key As String
    Dim path As String
    

    For y = Consts.BEIGN_YEAR To currYear
        For i = 0 To UBound(Consts.INDICES) - 1
            
            key = Config.IndexKey(Consts.INDICES(i), y)

            path = Config.SheetPath(key)
            
            Debug.Print path

            DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, key, path, True, "savedrecs!"
        Next i
    Next y

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



