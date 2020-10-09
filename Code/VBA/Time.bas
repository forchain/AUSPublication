Attribute VB_Name = "Time"
Option Compare Database
Option Explicit

Function currYear() As Integer

    currYear = Year(Date)
End Function

Function GetYear(PubYear, AccDate As Variant) As Integer
    Dim sAccDate As String
    PubYear = Trim(PubYear)
    AccDate = Trim(AccDate)
    If PubYear = "" Or IsNull(PubYear) Then

        Dim aAccDate() As String
        aAccDate = Split(AccDate)
        sAccDate = aAccDate(UBound(aAccDate))
    
    Else
        sAccDate = PubYear
    End If
    GetYear = CInt(sAccDate)
End Function

