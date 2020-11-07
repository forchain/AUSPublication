Attribute VB_Name = "Score"

Option Compare Database
Option Explicit



'iFalcC:  faculty Count
'iAuthC: author Count
'iPapInd: paper index
'iCurrInd: current index

Public Function CalcScore(ByVal IsStudent As Boolean, iPapInd As Integer, iCurrInd As Integer, iFacC As Integer, iAuthC As Integer) As Double

    If iAuthC = 0 Then
        'Debug.Print "[Error]CalcScore zero"
        Exit Function
    End If

    
    Dim dScore As Double
    dScore = 0#

    If iFacC = 0 Then                            ' without falcuty
        If IsStudent Then
            dScore = 1 / iAuthC
        End If
    Else                                         ' with faculty
        If Not IsStudent Then
            dScore = 1 / iFacC
        End If

    End If
    
    If iCurrInd = 0 Or iPapInd = iCurrInd Then
        CalcScore = FormatNumber(dScore, 2)
    Else
        CalcScore = 0#
    End If

End Function

