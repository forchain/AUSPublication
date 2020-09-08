Attribute VB_Name = "Author"
Option Compare Database
Option Explicit

Public Function ExtractOutDepID(Dep As String) As Integer
    Dim aToken
    aToken = Split(Dep, "-")
    ExtractOutDepID = CInt(Trim(aToken(0)))
End Function

Public Function ExtractOutCollName(Dep As String) As String
'Debug.Print "[Debug]ExtractOutCollName Dep:" & Dep
    Dim aToken() As String
    aToken = Split(Dep, "-")
    
    ExtractOutCollName = ExtractInCollName(aToken(1))
End Function

Public Function ExtractOutDepName(Dep As String) As String

'Debug.Print "[Debug]ExtractOutDepName Dep:" & Dep
    Dim aToken() As String
    aToken = Split(Dep, "-")
    
    ExtractOutDepName = ExtractInDepName(aToken(1))
End Function

Public Function ExtractInCollName(Dep As String) As String
    Dim aToken
    
    aToken = Split(Dep, "/")
    
    ExtractInCollName = Trim(aToken(0))
End Function

Public Function ExtractInDepName(Dep As String) As String
    Dim aToken
    
    aToken = Split(Dep, "/")
    
    ExtractInDepName = Trim(aToken(1))
End Function

Public Function GetAuthorName(Name As String) As String
    Dim aToken
    aToken = Split(Name, " ")
    
    Dim sFirstName As String
    sFirstName = Trim(aToken(0))
    
    Dim sLastName As String
    sLastName = Trim(aToken(UBound(aToken)))
    
    GetAuthorName = sFirstName + " " + sLastName
    
End Function

Public Function FixTitle(Title As String) As String
    Dim aToken() As String
    Dim sTitle As String
    
    sTitle = Replace(Title, "-Prof", " Prof")
    
    aToken = Split(sTitle, "-")
    
    sTitle = Trim(aToken(0))

    
    
    aToken = Split(sTitle, " of")
    
    sTitle = Trim(aToken(0))
    
    '    aToken = Split(sTitle, ".")
    '
    '    sTitle = Trim(aToken(0))
    
    aToken = Split(sTitle, "Prof")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Professor"
    End If
    
    aToken = Split(sTitle, "Inst")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Instructor"
    End If
    
    
    aToken = Split(sTitle, "Lecturer")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Lecturer"
    End If
    
    aToken = Split(sTitle, "Dean")
    
    sTitle = Trim(aToken(0))
    If UBound(aToken) > 0 Then
        sTitle = sTitle + " Dean"
    End If
    
    
    FixTitle = Trim(sTitle)

End Function





