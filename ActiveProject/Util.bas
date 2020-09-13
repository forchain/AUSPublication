Attribute VB_Name = "Util"
Option Compare Database
Option Explicit

 
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

