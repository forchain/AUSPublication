Attribute VB_Name = "AUS"
Option Compare Database
Option Explicit

Public Function ExtractAuthorsFromAddrs(Addrs As String) As String()
    
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    Dim sAuthors As String
    
    iEndPos = InStr(Addrs, "] Amer Univ Sharjah")
    If iEndPos = 0 Then
        Debug.Print "[Error]ExtractAuthorsFromAddrs No authors;Addrs:" & Addrs
        ExtractAuthorsFromAddrs = aAuthor
        
        Exit Function
    End If
    
    iStartPos = InStrRev(Addrs, "[", iEndPos)

    sAuthors = Mid(Addrs, iStartPos + 1, iEndPos - iStartPos - 1)
    
    aAuthor = Split(sAuthors, "; ")
  
         
    ExtractAuthorsFromAddrs = aAuthor
End Function

Public Function ExtractAuthorsFromIDs(IDs As String) As String()

    If IDs = "" Then
        Debug.Print "[Error]ExtractAuthorsFromIDs empty; IDs:" & IDs
        Exit Function
    End If
    
    Dim aAuthor() As String
    
    aAuthor = Split(IDs, ";")
    Dim i As Integer

    Dim a, Name As String
    For i = 0 To UBound(aAuthor)
        a = aAuthor(i)
        If a <> "" Then
            Name = Split(a, "/")(0)
            aAuthor(i) = Trim(Name)
        End If
    Next i
    
    ExtractAuthorsFromIDs = aAuthor
    
End Function

Public Function ExtractAuthorsText(Addrs As String) As String
    
    Dim iEndPos As Integer
    Dim iStartPos As Integer
    
    Dim aAuthor() As String
    Dim sAuthors As String
    
    iEndPos = InStr(Addrs, "] Amer Univ Sharjah")
    If iEndPos = 0 Then
        Debug.Print "[Error]ExtractAuthorsText No authors;Addrs:" & Addrs
        ExtractAuthorsText = ""
        
        Exit Function
    End If
    
    iStartPos = InStrRev(Addrs, "[", iEndPos)

    sAuthors = Mid(Addrs, iStartPos + 1, iEndPos - iStartPos - 1)
    
    ExtractAuthorsText = sAuthors

End Function

Public Function ExtractAuthors(Addrs As String) As String()
    
    Dim sAuthors, aAuthor() As String

    sAuthors = ExtractAuthorsText(Addrs)
    
    aAuthor = Split(sAuthors, "; ")
  
         
    ExtractAuthors = aAuthor
End Function

Public Function CountAuthors(Addrs As String) As Integer

    Dim aAuthor() As String
    aAuthor = ExtractAuthors(Addrs)
    If (Not Not aAuthor) = 0 Then
        Debug.Print "[Error]CountAuthors Addrs:" & Addrs
        CountAuthors = 0
        Exit Function
    End If
    CountAuthors = UBound(aAuthor) + 1
    'Debug.Print CountAuthors
End Function

Public Function SerializeAuthorNames(Addrs As String) As String

    Dim aAuthor() As String
    aAuthor = ExtractAuthors(Addrs)
    If UBound(aAuthor) = -1 Then
        Debug.Print "[Error]SerializeAuthorNames No authors;Addrs:" & Addrs
        Exit Function
    End If

    Dim names As String
    Dim i As Integer
    names = FixName(aAuthor(0))
    For i = 1 To UBound(aAuthor)
        names = names & ";" & FixName(aAuthor(i))
    Next i

    
    SerializeAuthorNames = names

End Function

Public Function FixName(FullName As String) As String
    If FullName = "" Then
        Debug.Print "[Error]FixName No Full name"
        FixName = FullName
        Exit Function
    End If

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
    
    sLastName = Trim(aFullName(0))
    If sLastName = "" Then
        Debug.Print "[Error]FixName No Last name; FullName:" & FullName
        FixName = FullName
        Exit Function
    End If
    
    sLastName = Split(sLastName, " ")(0)

    FixName = sFirstName + " " + sLastName
    'Debug.Print FixName
End Function

Public Function GetAbbrName(AuthorName As Variant) As String

    If IsNull(AuthorName) Then
        GetAbbrName = ""
        Exit Function
    End If
    Dim sFirstName, sLastName As String
    sFirstName = Left(AuthorName, 1) + "."
    sLastName = Split(AuthorName, " ")(1)

    GetAbbrName = sFirstName & " " & sLastName
End Function

Public Function FixNameWithIDs(Abbr As String, IDs As String) As String

    '    If Abbr = "W. Abuzaid" Then
    '        Debug.Print "[Debug]FixNameWithIDs Abbr:" & Abbr & ", IDs:" & IDs
    '    End If
    
    If Mid(Abbr, 2, 1) <> "." Then
        FixNameWithIDs = Abbr
        Exit Function
    End If

    If IDs = "" Then
        FixNameWithIDs = Abbr
        Debug.Print "[Warn]FixNameWithIDs empty; Abbr:" & Abbr

        Exit Function
    End If
    
    Dim aAuthor() As String
    aAuthor = ExtractAuthorsFromIDs(IDs)
    Dim a As String
    Dim i As Integer
    For i = 0 To UBound(aAuthor)
        a = aAuthor(i)
        If a <> "" Then
            a = FixName(a)
            If (Mid(a, 2, 1) <> ".") And (Left(a, 1) = Left(Abbr, 1)) And (Left(a, 1) <> ",") Then
                Dim lastName As String
                lastName = Split(Abbr, " ")(1)
                If InStr(a, lastName) <> 0 Then
                    'Debug.Print "[Trace]FixNameWithIDs fixed; Abbr:" & Abbr & ", IDs:" & IDs
                    FixNameWithIDs = a
                    Exit Function
                End If
            End If
        End If
    Next i
    
    Debug.Print "[Warn]FixNameWithIDs unfixed; Abbr:" & Abbr & ", IDs:" & IDs
    FixNameWithIDs = Abbr

End Function

Public Function SelectAuthor(Addrs As String, Order As Integer, ResearcherIDs As String, ORCIDs As String) As Variant

    If Order > 9 Or Order < 1 Then
        Debug.Print "SelectAuthor error, Order: " & CStr(Order)
        SelectAuthor = Null
        Exit Function
    End If



    Dim aAuthor() As String

    aAuthor = ExtractAuthorsFromAddrs(Addrs)
    If (Not Not aAuthor) = 0 Then
        Debug.Print "[Error]SelectAuthor No Authors; Order: " & CStr(Order) & ", Addrs:" & Addrs

        SelectAuthor = Null
        Exit Function
    End If
    Dim iIndex As Integer
    iIndex = Order - 1

    If UBound(aAuthor) >= 9 Then
        Debug.Print "SelectAuthor warning, UBound >= " & CStr(UBound(aAuthor))
    End If


    If iIndex > UBound(aAuthor) Then
        'Debug.Print "SelectAuthor error, iIndex > " & CStr(UBound(aAuthor))
        SelectAuthor = Null
        Exit Function
    End If

    Dim fixedName As String
    fixedName = FixName(aAuthor(iIndex))
    
    fixedName = FixNameWithIDs(fixedName, ResearcherIDs)
    fixedName = FixNameWithIDs(fixedName, ORCIDs)

    SelectAuthor = fixedName
    'Debug.Print "[Trace]SelectAuthor fixedName:" & fixedName

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

