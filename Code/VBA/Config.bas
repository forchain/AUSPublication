Attribute VB_Name = "Config"

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


Public Property Get Setting(Key As String) As Variant

    Dim rsSetting As Recordset
    Set rsSetting = CurrentDb.OpenRecordset("Setting", dbOpenSnapshot)

    Setting = rsSetting(Key).Value
    rsSetting.Close
End Property

