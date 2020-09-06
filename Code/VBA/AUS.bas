Option Compare Database
Option Explicit

Public Function ExtractAUSAuthors(sAddrs As String) As String()
    
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    Dim sAuthors As String
    
    iEndPos = InStr(sAddrs, "] Amer Univ Sharjah")
    If iEndPos = 0 Then
        Debug.Print "Cannot find authors in " & sAddrs
        ExtractAUSAuthors = aAuthor
        
        Exit Function
    End If
    
    iStartPos = InStrRev(sAddrs, "[", iEndPos)

    sAuthors = Mid(sAddrs, iStartPos + 1, iEndPos - iStartPos - 1)
    
    aAuthor = Split(sAuthors, "; ")
  
    Dim i As Integer
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


Public Function ExtractAuthors(Addrs As String) As String()
    
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    Dim sAuthors As String
    
    iEndPos = InStr(Addrs, "] Amer Univ Sharjah")
    If iEndPos = 0 Then
        Debug.Print "Cannot find authors in " & Addrs
        ExtractAuthors = aAuthor
        
        Exit Function
    End If
    
    iStartPos = InStrRev(Addrs, "[", iEndPos)

    sAuthors = Mid(Addrs, iStartPos + 1, iEndPos - iStartPos - 1)
    
    aAuthor = Split(sAuthors, "; ")
  
         
    ExtractAuthors = aAuthor
End Function



Public Function CountAuthors(Addrs As String) As Integer

    Dim aAuthors() As String
    aAuthors = ExtractAuthors(Addrs)
    CountAuthors = UBound(aAuthors) + 1
    'Debug.Print CountAuthors
End Function

Public Function FixName(FullName As String) As String

    Dim aFullName As Variant
    
    aFullName = Split(Trim(FullName), ",")

    If UBound(aFullName) = 0 Then
        Debug.Print "FixName warning, " & FullName
        FixName = FullName
        Debug.Print FixName
        Exit Function
    End If
    ' WoS naming style: Last Name, First Name
    Dim sFirstName, sLastName As String
    sFirstName = Split(Trim(aFullName(1)), " ")(0)
    sLastName = Split(Trim(aFullName(0)), " ")(0)

    FixName = sFirstName + " " + sLastName
    'Debug.Print FixName
End Function

Public Function SelectAuthor(Addrs As String, Order As String) As Variant

    If Order > 9 Or Order < 1 Then
        Debug.Print "SelectAuthor error, Order: " & CStr(Order)
        SelectAuthor = Null
        Exit Function
    End If



    Dim aAuthors() As String

    aAuthors = ExtractAuthors(Addrs)

    Dim iIndex As Integer
    iIndex = Order - 1

    If UBound(aAuthors) >= 9 Then
        Debug.Print "SelectAuthor warning, UBound >= " & CStr(UBound(aAuthors))
    End If


    If iIndex > UBound(aAuthors) Then
        'Debug.Print "SelectAuthor error, iIndex > " & CStr(UBound(aAuthors))
        SelectAuthor = Null
        Exit Function
    End If

    SelectAuthor = FixName(aAuthors(iIndex))

    Debug.Print SelectAuthor

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

