Attribute VB_Name = "Match"
Option Compare Database
Option Explicit

Public Function IsMatched(PaperFirstName, PaperMiddleName, PaperMiddleInitial, _
                          AuthorID, AuthorFirstName, AuthorFirstInitial, AuthorMiddleName, AuthorMiddleInitial, _
                          FirstNameRequired, FirstNameMatched, _
                          MiddleNameRequired, MiddleNameMatched, _
                          MiddleInitialRequired, MiddleInitialMatched As Variant) As Boolean

    IsMatched = False

    If IsNull(AuthorID) Then
        Exit Function
    End If

    ' Must exist AuthorFirstName
    If FirstNameRequired Then
        If IsNull(AuthorFirstName) Then
            Exit Function
        End If
    End If

    ' Must match AuthorFirstName
    If FirstNameMatched Then
        If Not IsNull(PaperFirstName) And Not IsNull(AuthorFirstName) Then
            If PaperFirstName <> AuthorFirstName Then
                Exit Function
            End If
        End If
    End If

    ' Must exist AuthorMiddleName
    If MiddleNameRequired Then
        If IsNull(AuthorMiddleName) Then
            Exit Function
        End If
    End If

    ' Must  match AuthorMiddleName
    Dim i As Integer
    
    If MiddleNameMatched Then
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

    ' Must exist AuthorMiddleInitial
    If MiddleInitialRequired Then
        If IsNull(AuthorMiddleInitial) Then
            Exit Function
        End If
    End If

    ' Must  match AuthorMiddleInitial
    If MiddleInitialMatched Then
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

