Option Compare Database
Option Explicit

Public Function ExtractAUSAuthors(sAddrs As String) As String()
    
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    
    iEndPos = InStr(sAddrs, "] Amer Univ Sharjah")
    If iEndPos = 0 Then
        Debug.Print "Cannot find authors in " & sAddrs
        ExtractAUSAuthors = aAuthor
        
        Exit Function
    End If
    
    iStartPos = InStrRev(sAddrs, "[", iEndPos)
    sAuthors = Mid(sAddrs, iStartPos + 1, iEndPos - iStartPos - 1)
    
    aAuthor = Split(sAuthors, "; ")
  
    For i = 0 To UBound(aAuthor)
    
        aAuthor(i) = Trim(aAuthor(i))
        If Right(aAuthor(i), 1) = "." And Right(aAuthor(i), 4) <> "," Then
            'Debug.Print aAuthor(i)
            aAuthor(i) = Mid(aAuthor(i), 1, Len(aAuthor(i)) - 3)
            'Debug.Print aAuthor(i)
        End If
    Next
         
    ExtractAUSAuthors = aAuthor

End Function

' LastName, FirstName
Public Function GetFirstName(sFullName) As String
    aFullName = Split(sFullName, ",")
    If UBound(aFullName) = 0 Then
        Debug.Print "GetFirstName failed, " & sFullName
        GetFirstName = ""
        Exit Function
    End If
    GetFirstName = Trim(aFullName(1))
End Function

Public Function GetLastName(sFullName) As String
    aFullName = Split(sFullName, ",")

    GetLastName = Trim(aFullName(0))
End Function

