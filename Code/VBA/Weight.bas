Attribute VB_Name = "Weight"

Option Compare Database
Option Explicit



'iFalcC:  faculty Count
'iAuthC: author Count
'iPapInd: paper index
'iCurrInd: current index

Public Function CalcScore(iID As Variant, iPapInd As Integer, iCurrInd As Integer, iFacC As Integer, iAuthC As Integer) As Double

    If iAuthC = 0 Then
        'Debug.Print "[Error]CalcScore zero"
        Exit Function
    End If


    Dim bIsFac As Byte

    If IsNull(iID) Then
        bIsFac = False
    Else
        bIsFac = True
    End If
    
    
    
    Dim dScore As Double
    dScore = 0#

    If iFacC = 0 Then                            ' without falcuty
        If Not bIsFac Then
            dScore = 1 / iAuthC
        End If
    Else                                         ' with faculty
        If bIsFac Then
            dScore = 1 / iFacC
        End If

    End If
    
    If iCurrInd = 0 Or iPapInd = iCurrInd Then
        CalcScore = FormatNumber(dScore, 2)
    Else
        CalcScore = 0#
    End If

End Function

Public Function CalcMatchingScore(WeightFirstName, WeightMiddleName, WeightMiddleInitial, _
                                  Code, AuthorFirstName, AuthorFirstInitial, AuthorMiddleName, AuthorMiddleInitial As Variant) As Integer
    Dim lScore As Integer

    lScore = 0
    
    If Not IsNull(Code) Then
        lScore = 2 ^ 0
    Else
        CalcMatchingScore = lScore
        Exit Function
    End If

    If Not IsNull(WeightFirstName) And Not IsNull(AuthorFirstName) Then
        lScore = lScore + 2 ^ 1
        
        If WeightFirstName <> AuthorFirstName Then
            CalcMatchingScore = lScore
            Exit Function
        Else
            lScore = lScore + 2 ^ 2
        End If

    End If

    Dim i As Integer
    If Not IsNull(WeightMiddleName) And Not IsNull(AuthorMiddleName) Then
        lScore = lScore + 2 ^ 3
        
        Dim sWeightMiddleName As String
        Dim vWeightMiddleName As Variant
        vWeightMiddleName = Split(WeightMiddleName)
        For i = 0 To UBound(vWeightMiddleName)
            sWeightMiddleName = vWeightMiddleName(i)
            If InStr(AuthorMiddleName, sWeightMiddleName) = 0 Then
                CalcMatchingScore = lScore
                Exit Function
            End If
        Next
        
        lScore = lScore + 2 ^ 4
 
    End If

    If Not IsNull(WeightMiddleInitial) And Not IsNull(AuthorMiddleInitial) Then
        lScore = lScore + 2 ^ 5

        Dim sWeightMiddleInitial As String
        Dim vWeightMiddleInitial As Variant
        vWeightMiddleInitial = Split(WeightMiddleInitial)
        For i = 0 To UBound(vWeightMiddleInitial)
            sWeightMiddleInitial = vWeightMiddleInitial(i)
            If InStr(AuthorMiddleInitial, sWeightMiddleInitial) = 0 Then
                CalcMatchingScore = lScore
                Exit Function
            End If
        Next
        lScore = lScore + 2 ^ 6
    End If
    CalcMatchingScore = lScore
End Function




























