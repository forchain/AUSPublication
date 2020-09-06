Option Compare Database
Option Explicit

Public Property Get SettingPath() As String
    SettingPath = CurrentProject.path + Consts.SETTINGS_FILE
End Property

Public Property Get IndexKey(ByVal Index As String, Year As Integer) As String

    IndexKey = Index + "-" + CStr(Year)
    
End Property

Public Property Get SheetPath(key As String) As String

    SheetPath = CurrentProject.path + Consts.SHEETS_DIR + Val(Consts.SECTION_INDEX, key)
    
End Property


Public Property Get Val(Section As String, key As String) As String
    Val = Word.System.PrivateProfileString(SettingPath, Section, key)
End Property

Public Property Let Val(Section As String, key As String, Value As String)
    Word.System.PrivateProfileString(SettingPath, Section, key) = Value
End Property

