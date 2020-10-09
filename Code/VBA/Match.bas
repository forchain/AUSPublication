Attribute VB_Name = "Match"
Option Compare Database
Option Explicit

Public Function IsMatched(PaperFirstName, PaperMiddleName, PaperMiddleInitial, _
                                  AuthorID, AuthorFirstName, AuthorFirstInitial, AuthorMiddleName, AuthorMiddleInitial, 
                                                                  FirstNameRequired ,
    FirstNameMatched ,
    MiddleNameRequired ,
    MiddleNameMatched ,
    MiddleInitialRequired ,
    MiddleInitialMatched  
                                  
                                  As Variant) As Integer

    IsMatched = False

    If  IsNull(AuthorID) Then
    Exit Function
    end if

' Must exist AuthorFirstName
    If  FirstNameRequired Then
        if IsNull(AuthorFirstName) Then
        exit Function
        end if
    end if

' Must match AuthorFirstName
    If FirstNameMatched Then
    If Not IsNull(PaperFirstName) And Not IsNull(AuthorFirstName) Then
        If PaperFirstName <> AuthorFirstName Then
            Exit Function
        End If
    End If
    end if

' Must exist AuthorMiddleName
    If MiddleNameRequired  Then
        if IsNull(AuthorMiddleName) Then
        exit Function
        end if
    end if

' Must  match AuthorMiddleName
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
    end if

' Must exist AuthorMiddleInitial
    If MiddleInitialRequired Then
        if IsNull(AuthorMiddleInitial) Then
        exit Function
        end if
    end if

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
    end if

    IsMatched = True
End Function


Public Function CalcPoints(Points, Condition as Integer) as Boolean

if Points& 2^0
    If True Then
        
    Else
        
    End If


    CalcPoints=False

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