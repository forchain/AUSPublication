Attribute VB_Name = "Log"

Option Compare Database
Option Explicit

Private Function Out(Tag, Func, Reason As String, ParamArray Vars() As Variant) As String

    Dim vVars() As Variant
    vVars = Vars(0)
    Dim sMes As String
 
    Dim sKey, sVal, sVars As String
    sVars = ";"
    Dim i As Integer
    For i = 0 To UBound(vVars) Step 2
        sKey = CStr(vVars(i))
        sVal = CStr(vVars(i + 1))

        sVars = sVars & " " & sKey & ":" & sVal
    Next i

    sMes = "[" & Tag & "]" & Func & " " & Reason & sVars


    Debug.Print sMes
    Out = sMes

End Function

Public Sub D(Func, Reason As String, ParamArray Vars() As Variant)
    Dim sLog As String
    sLog = Out("D", Func, Reason, Vars)
End Sub

Public Sub i(Func, Reason As String, ParamArray Vars() As Variant)
    Dim sLog As String
    sLog = Out("I", Func, Reason, Vars)
End Sub

Public Sub W(Func, Reason As String, ParamArray Vars() As Variant)
    Dim sLog As String
    sLog = Out("W", Func, Reason, Vars)
End Sub

Public Sub T(Func, Reason As String, ParamArray Vars() As Variant)
    Dim sLog As String
    sLog = Out("T", Func, Reason, Vars)
End Sub

Public Sub E(Func, Reason As String, ParamArray Vars() As Variant)
    Dim sLog As String
    sLog = Out("E", Func, Reason, Vars)
    
    Dim rResult As VbMsgBoxResult
    rResult = MsgBox(Reason, vbYes, Func)
End Sub


