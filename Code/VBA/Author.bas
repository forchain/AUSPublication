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

Public Function GetAuthorName(ByVal Name As String) As Variant
    Dim aToken
    aToken = Split(Name, " ")
    
    Dim sFirstName As String
    sFirstName = Trim(aToken(0))
    
    Dim sLastName As String
    sLastName = Trim(aToken(UBound(aToken)))
    
    GetAuthorName = sFirstName + " " + sLastName
    
End Function

Public Function GetAuthorFirstName(ByVal FullName) As Variant
    Dim sFN As String
    Dim aToken As Variant
    aToken = Split(FullName, " ")
    
    sFN = aToken(0)
    
    If Len(sFN) = 2 And Right(sFN, 1) = "." Then
        GetAuthorFirstName = Null
        Exit Function
    End If
    
    GetAuthorFirstName = Trim(aToken(0))
End Function

Public Function GetAuthorFirstInitial(ByVal FullName) As Variant
    GetAuthorFirstInitial = Left(FullName, 1) + "."
End Function

Public Function GetAuthorLastName(ByVal FullName) As Variant
    Dim aToken() As String
    aToken = Split(FullName, " ")
    
    GetAuthorLastName = Trim(aToken(UBound(aToken)))
End Function

Public Function GetAuthorMiddleName(ByVal FullName As String) As Variant
    Dim aFullName() As String
    aFullName = Split(FullName, " ")
    
    If UBound(aFullName) < 2 Then
        GetAuthorMiddleName = Null
        Exit Function
    End If

    Dim iLen As Integer
    ' deduce FN, LN
    iLen = UBound(aFullName) - 1

    Dim aMiddleName() As String
    ReDim aMiddleName(iLen)

    Dim i As Integer
    For i = 1 To UBound(aFullName) - 1
        If Len(aFullName(i)) = 1 Then
            GetAuthorMiddleName = Null
            Exit Function
        End If
        aMiddleName(i - 1) = (aFullName(i))
    Next
    
    GetAuthorMiddleName = Trim(Join(aMiddleName, " "))
End Function

Public Function GetAuthorMiddleInitial(ByVal FullName As String) As Variant
    Dim aFullName() As String
    aFullName = Split(FullName, " ")
    
    If UBound(aFullName) < 2 Then
        GetAuthorMiddleInitial = Null
        Exit Function
    End If

    Dim iLen As Integer
    ' deduce FN, LN
    iLen = UBound(aFullName) - 1

    Dim aMiddleName() As String
    ReDim aMiddleName(iLen)

    Dim i As Integer
    For i = 1 To UBound(aFullName) - 1
        aMiddleName(i - 1) = Left(aFullName(i), 1) + "."
    Next
    
    GetAuthorMiddleInitial = Trim(Join(aMiddleName, " "))
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

