Option Compare Database
 
Private lngRowNumber As Long
Private colPrimaryKeys As VBA.Collection

Private dblRunningSum As Double
Private colKeyRS As VBA.Collection

Public Function ResetRowNumber() As Boolean
    Set colPrimaryKeys = New VBA.Collection
    lngRowNumber = 0

    Set colKeyRS = New VBA.Collection
    dblRunningSum = 0

    ResetRowNumber = True
End Function

Public Function RowNumber(UniqueKeyVariant As Variant) As Long
    Dim lngTemp As Long

    On Error Resume Next
    lngTemp = colPrimaryKeys(CStr(UniqueKeyVariant))
    If Err.Number Then
        lngRowNumber = lngRowNumber + 1
        colPrimaryKeys.Add lngRowNumber, CStr(UniqueKeyVariant)
        lngTemp = lngRowNumber
    End If

    RowNumber = lngTemp
End Function

Public Function RunningSum(UniqueKeyVariant As Variant, varNum As Variant) As Double
    Dim dblTemp As Double

    On Error Resume Next
    dblTemp = colKeyRS(CStr(UniqueKeyVariant))
    If Err.Number Then
        dblRunningSum = dblRunningSum + CDbl(varNum)
        colPrimaryKeys.Add dblRunningSum, CStr(UniqueKeyVariant)
        dblTemp = dblRunningSum
    End If

    RunningSum = dblTemp
End Function

Public Function TokenNum(strValues As String, strDelim As String) As Double

    Dim arSplit As Variant

    arSplit = Split(strValues, strDelim)

    TokenNum = UBound(arSplit) + 1

End Function

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

