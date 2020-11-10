Attribute VB_Name = "Match"
Option Compare Database
Option Explicit

' Minimum match: last name and first initial
Public Function IsMinimumMatched(ByVal AuthorID As Variant) As Boolean
    IsMinimumMatched = False
    If IsNull(AuthorID) Then
        Exit Function
    End If
    IsMinimumMatched = True
End Function

Public Function IsFirstNameMatched(ByVal FirstNameCheck As Boolean, PaperFirstName, AuthorFirstName As Variant) As Boolean

    IsFirstNameMatched = False
    ' Must exist AuthorFirstName
    ' No need to check author first name since it cannot be empty
    If FirstNameCheck Then
        If IsNull(AuthorFirstName) Or IsNull(PaperFirstName) Then
            Exit Function
        End If
    End If

    ' Must match AuthorFirstName
    
    If Not IsNull(PaperFirstName) And Not IsNull(AuthorFirstName) Then
        If PaperFirstName <> AuthorFirstName Then
            Exit Function
        End If
    End If

    IsFirstNameMatched = True


End Function

Public Function IsMiddleNameMatched(ByVal MiddleNameCheck As Boolean, PaperMiddleName, AuthorMiddleName As Variant) As Boolean
    IsMiddleNameMatched = False
    ' Must exist AuthorMiddleName
    If MiddleNameCheck Then
        If (IsNull(AuthorMiddleName) And Not IsNull(PaperMiddleName)) Or (Not IsNull(AuthorMiddleName) And IsNull(PaperMiddleName)) Then
            Exit Function
        End If
    End If

    ' Must  match AuthorMiddleName
    Dim i As Integer
    

    If Not IsNull(PaperMiddleName) And Not IsNull(AuthorMiddleName) Then
        Dim sPaperMiddleName As String
        Dim vPaperMiddleName As Variant
        vPaperMiddleName = Split(PaperMiddleName)
        For i = 0 To UBound(vPaperMiddleName)
            sPaperMiddleName = vPaperMiddleName(i)
            If InStr(AuthorMiddleName, sPaperMiddleName) = 0 Then
                Exit Function
            End If
        Next
    End If
    IsMiddleNameMatched = True
End Function

Public Function IsMiddleInitialMatched(ByVal MiddleInitialCheck As Boolean, PaperMiddleInitial, AuthorMiddleInitial As Variant) As Boolean
    IsMiddleInitialMatched = False
    ' Must exist AuthorMiddleInitial
    If MiddleInitialCheck Then
        If (IsNull(AuthorMiddleInitial) And Not IsNull(PaperMiddleInitial)) Or (Not IsNull(AuthorMiddleInitial) And IsNull(PaperMiddleInitial)) Then
            Exit Function
        End If
    End If

    ' Must  match AuthorMiddleInitial

    Dim i As Integer

    If Not IsNull(PaperMiddleInitial) And Not IsNull(AuthorMiddleInitial) Then
        Dim sPaperMiddleInitial As String
        Dim vPaperMiddleInitial As Variant
        vPaperMiddleInitial = Split(PaperMiddleInitial)
        For i = 0 To UBound(vPaperMiddleInitial)
            sPaperMiddleInitial = vPaperMiddleInitial(i)
            If InStr(AuthorMiddleInitial, sPaperMiddleInitial) = 0 Then
                Exit Function
            End If
        Next
    End If
    
    IsMiddleInitialMatched = True

End Function

Public Function IsMatched(ByVal PaperFirstName, PaperMiddleName, PaperMiddleInitial, _
                          AuthorID, AuthorFirstName, AuthorMiddleName, AuthorMiddleInitial As Variant, _
                          FirstNameCheck, MiddleNameCheck, MiddleInitialCheck As Boolean) As Boolean

    IsMatched = False
    
    
    If Not IsMinimumMatched(AuthorID) Then
        Exit Function
    End If
    
    If Not IsFirstNameMatched(FirstNameCheck, PaperFirstName, AuthorFirstName) Then
        Exit Function
    End If

    If Not IsMiddleNameMatched(MiddleNameCheck, PaperMiddleName, AuthorMiddleName) Then
        Exit Function
    End If

    If Not IsMiddleInitialMatched(MiddleInitialCheck, PaperMiddleInitial, AuthorMiddleInitial) Then
        Exit Function
    End If

    IsMatched = True
End Function

Public Function CalcPoints(PaperFirstName, PaperMiddleName, PaperMiddleInitial, _
                           AuthorCode, AuthorFirstName, AuthorFirstInitial, AuthorMiddleName, AuthorMiddleInitial As Variant) As Integer
    Dim lScore As Integer

    lScore = 0
    
    If Not IsNull(AuthorCode) Then
        lScore = 2 ^ 0
    Else
        CalcMatchingScore = lScore
        Exit Function
    End If

    If Not IsNull(PaperFirstName) And Not IsNull(AuthorFirstName) Then
        lScore = lScore + 2 ^ 1
        
        If PaperFirstName <> AuthorFirstName Then
            CalcMatchingScore = lScore
            Exit Function
        Else
            lScore = lScore + 2 ^ 2
        End If

    End If

    Dim i As Integer
    If Not IsNull(PaperMiddleName) And Not IsNull(AuthorMiddleName) Then
        lScore = lScore + 2 ^ 3
        
        Dim sPaperMiddleName As String
        Dim vPaperMiddleName As Variant
        vPaperMiddleName = Split(PaperMiddleName)
        For i = 0 To UBound(vPaperMiddleName)
            sPaperMiddleName = vPaperMiddleName(i)
            If InStr(AuthorMiddleName, sPaperMiddleName) = 0 Then
                CalcMatchingScore = lScore
                Exit Function
            End If
        Next
        
        lScore = lScore + 2 ^ 4
 
    End If

    If Not IsNull(PaperMiddleInitial) And Not IsNull(AuthorMiddleInitial) Then
        lScore = lScore + 2 ^ 5

        Dim sPaperMiddleInitial As String
        Dim vPaperMiddleInitial As Variant
        vPaperMiddleInitial = Split(PaperMiddleInitial)
        For i = 0 To UBound(vPaperMiddleInitial)
            sPaperMiddleInitial = vPaperMiddleInitial(i)
            If InStr(AuthorMiddleInitial, sPaperMiddleInitial) = 0 Then
                CalcMatchingScore = lScore
                Exit Function
            End If
        Next
        lScore = lScore + 2 ^ 6
    End If
    CalcMatchingScore = lScore
End Function


